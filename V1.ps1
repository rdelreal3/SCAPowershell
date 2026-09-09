# MSSA Employee Onboarding, Identity Access, and Access Audit Tool
# Workflow: Workday Export -> Power Automate -> PowerShell -> Active Directory -> Reports
# Phase 1: Receive employee information.
# Phase 2: Validate employee information and log exceptions.
# Phase 3: Verify the Domain Controller and collect CIM information.
# Phase 4: Check for duplicates and create the AD user.
# Phase 5: Determine and assign the correct OU.
# Phase 6: Assign role-based security groups.
# Phase 7: Perform a randomized read-only OU access audit.
# Phase 8: Export onboarding, exception, and audit reports.
# Phase 9: Display final status and handle errors.

# Power Automate passes the Workday file path. Individual fields and JSON support testing.
[CmdletBinding()]
param(
    [string]$EmployeeFilePath,
    [string]$EmployeeJson,
    [string]$FirstName,
    [string]$LastName,
    [string]$JobTitle,
    [string]$Department,
    [string]$Manager,
    [string]$Location,
    [string]$Email,
    [string]$EmployeeType,
    [string]$PhoneNumber,
    [string]$EmployeeId,
    [securestring]$TemporaryPassword,
    [int]$AuditSampleSize = 0,
    [switch]$DryRun
)

# Environment-specific values are kept here so students can change one section.
# FILL IN LATER: Replace every value marked with FILL_IN before using the script.
# Confirm these values with the Windows Server, Active Directory, and HR teams.
$script:Config = [ordered]@{
    DomainController = 'FILL_IN_DOMAIN_CONTROLLER_NAME'
    DomainName = 'contoso.com'
    ClientComputer = 'FILL_IN_CLIENT_COMPUTER_NAME'
    ServerComputer = 'FILL_IN_SERVER_COMPUTER_NAME'
    ReportFolder = 'FILL_IN_REPORT_FOLDER_PATH'
    LogFolder = 'FILL_IN_LOG_FOLDER_PATH'
    HrContact = 'FILL_IN_HR_ONBOARDING_DIRECTOR_NAME_AND_CONTACT'
    DefaultAuditSampleSize = 3 # FILL IN LATER if the audit sample size should change.
}
# Temporary local defaults are used only when report paths are still being gathered.
if ($script:Config.ReportFolder -like 'FILL_IN*') { $script:Config.ReportFolder = Join-Path (Get-Location) 'Reports' }
if ($script:Config.LogFolder -like 'FILL_IN*') { $script:Config.LogFolder = Join-Path (Get-Location) 'Logs' }
if ($AuditSampleSize -le 0) { $AuditSampleSize = $script:Config.DefaultAuditSampleSize }

# FILL IN LATER: Replace each placeholder with the real distinguished name.
# No fallback OU is used when an approved mapping is missing.
$script:OuMappings = @{
    'IT|Employee' = 'FILL_IN_IT_EMPLOYEE_OU_DN'
    'IT|Contractor' = 'FILL_IN_IT_CONTRACTOR_OU_DN'
    'Finance|Employee' = 'FILL_IN_FINANCE_EMPLOYEE_OU_DN'
    'Human Resources|Employee' = 'FILL_IN_HR_EMPLOYEE_OU_DN'
    'Sales|Employee' = 'FILL_IN_SALES_EMPLOYEE_OU_DN'
}
# FILL IN LATER: Replace group names with approved groups from Active Directory.
$script:GroupMappings = @{
    'Department:IT' = @('FILL_IN_IT_SECURITY_GROUP')
    'Department:Finance' = @('FILL_IN_FINANCE_SECURITY_GROUP')
    'Department:Human Resources' = @('FILL_IN_HR_SECURITY_GROUP')
    'Department:Sales' = @('FILL_IN_SALES_SECURITY_GROUP')
    'EmployeeType:Employee' = @('FILL_IN_EMPLOYEE_SECURITY_GROUP')
    'EmployeeType:Contractor' = @('FILL_IN_CONTRACTOR_SECURITY_GROUP')
    'EmployeeType:Intern' = @('FILL_IN_INTERN_SECURITY_GROUP')
    'JobTitle:Manager' = @('FILL_IN_MANAGER_SECURITY_GROUP')
}

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$script:State = [ordered]@{
    EmployeeName = 'Unknown employee'; Username = ''; AssignedOu = ''
    AssignedGroups = [System.Collections.ArrayList]::new()
    Errors = [System.Collections.ArrayList]::new()
    Status = 'Started'; StartedAt = Get-Date; CompletedAt = $null
}

# ============================================================================
# PHASE 1: RECEIVE EMPLOYEE INFORMATION
# Purpose: Read exactly one Workday CSV or JSON record. Import-Csv converts column
# headers into object properties. ConvertFrom-Json supports Power Automate JSON.
# Test-Path confirms the file exists. No administrator is prompted to invent HR data.
# ============================================================================
function Import-EmployeeRecord {
    param([string]$FilePath, [string]$Json)
    Write-Host '[PHASE 1] Receiving Workday employee information...'
    if ($FilePath) {
        if (-not (Test-Path -LiteralPath $FilePath -PathType Leaf)) { throw "File not found: $FilePath" }
        $extension = [IO.Path]::GetExtension($FilePath).ToLowerInvariant()
        if ($extension -eq '.csv') { $records = @(Import-Csv -LiteralPath $FilePath) }
        elseif ($extension -eq '.json') { $records = @(Get-Content -LiteralPath $FilePath -Raw | ConvertFrom-Json) }
        else { throw "Unsupported file type '$extension'. Use .csv or .json." }
        if ($records.Count -ne 1) { throw "Expected one Workday record, found $($records.Count)." }
        return $records[0]
    }
    if ($Json) { return ($Json | ConvertFrom-Json) }
    return [pscustomobject]@{
        FirstName = $FirstName; LastName = $LastName; JobTitle = $JobTitle
        Department = $Department; Manager = $Manager; Location = $Location
        Email = $Email; EmployeeType = $EmployeeType; PhoneNumber = $PhoneNumber
        EmployeeId = $EmployeeId
    }
}
# Presentation Notes:
# "Phase 1 receives the Workday record. Power Automate watches the folder and passes
# the new file path. Import-Csv creates an object from the column headers, allowing
# later commands to use readable properties. One record per run prevents accidental
# bulk processing, and Workday remains the source of truth."

# ============================================================================
# PHASE 2: VALIDATE EMPLOYEE INFORMATION AND LOG EXCEPTIONS
# Purpose: A hashtable pairs friendly names with values. The pipeline and Where-Object
# identify null, empty, or whitespace values. A daily CSV exception is written before
# the workflow stops, so administrators can open it in Excel.
# ============================================================================
function New-ExceptionRecord {
    param([string]$EmployeeName, [string]$MissingFields, [string]$ErrorDetails)
    New-Item -Path $script:Config.ReportFolder -ItemType Directory -Force | Out-Null
    $file = Join-Path $script:Config.ReportFolder ("OnboardingExceptions_{0}.csv" -f (Get-Date -Format 'yyyy-MM-dd'))
    $record = [pscustomobject]@{
        EmployeeName = $EmployeeName; MissingFields = $MissingFields
        DateTime = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'; Status = 'Validation Failed'
        RequiredAction = "HR must update Workday and re-export. Escalate to $($script:Config.HrContact)."
        ErrorDetails = $ErrorDetails
    }
    if (Test-Path -LiteralPath $file) { $record | Export-Csv $file -NoTypeInformation -Append }
    else { $record | Export-Csv $file -NoTypeInformation }
}
function Test-EmployeeRecord {
    param([Parameter(Mandatory)] [psobject]$Employee)
    Write-Host '[PHASE 2] Validating all required fields...'
    $fieldNames = [ordered]@{
        'First Name' = 'FirstName'; 'Last Name' = 'LastName'; 'Job Title' = 'JobTitle'
        Department = 'Department'; Manager = 'Manager'; Location = 'Location'
        Email = 'Email'; 'Employee Type' = 'EmployeeType'; 'Phone Number' = 'PhoneNumber'
    }
    $requiredFields = [ordered]@{}
    foreach ($name in $fieldNames.Keys) {
        $property = $Employee.PSObject.Properties[$fieldNames[$name]]
        $requiredFields[$name] = if ($property) { $property.Value } else { $null }
    }
    $missing = @($requiredFields.GetEnumerator() | Where-Object {
        [string]::IsNullOrWhiteSpace([string]$_.Value)
    } | Select-Object -ExpandProperty Key)
    if ($missing.Count -gt 0) {
        $name = "$(if ($Employee.PSObject.Properties['FirstName']) {$Employee.FirstName}) $(if ($Employee.PSObject.Properties['LastName']) {$Employee.LastName})".Trim()
        $details = "Missing required fields: $($missing -join ', ')"
        New-ExceptionRecord -EmployeeName $name -MissingFields ($missing -join ', ') -ErrorDetails $details
        throw "$details. HR must correct Workday and re-export. Contact $($script:Config.HrContact)."
    }
    if (-not $Employee.PSObject.Properties['EmployeeId']) { $Employee | Add-Member EmployeeId '' }
    Write-Host '[PHASE 2] Validation succeeded.'
}
# Presentation Notes:
# "Phase 2 is the safety gate. The readable hashtable and Where-Object pipeline identify
# every missing field, not just the first one. The exact fields are appended to a daily
# exception CSV. The script stops and directs HR to correct Workday; it never asks an
# administrator to manually supply missing HR information."

# ============================================================================
# PHASE 3: VERIFY DOMAIN CONTROLLER AND COLLECT CIM INFORMATION
# Purpose: Test-Connection checks DC1 before changes. Get-CimInstance collects useful
# operating-system data, and Select-Object controls the returned fields.
# ============================================================================
function Test-DomainController {
    param([string]$ComputerName)
    Write-Host "[PHASE 3] Checking $ComputerName with Test-Connection..."
    if (-not (Test-Connection -ComputerName $ComputerName -Count 2 -Quiet)) { throw "Cannot reach $ComputerName. Onboarding stopped." }
    $info = Get-CimInstance Win32_OperatingSystem -ComputerName $ComputerName -ErrorAction Stop |
        Select-Object @{Name='ComputerName';Expression={$_.CSName}}, Caption, Version, BuildNumber, LastBootUpTime
    Write-Host "[PHASE 3] $($info.ComputerName): $($info.Caption), build $($info.BuildNumber)"
    return $info
}
# Presentation Notes:
# "Phase 3 checks infrastructure before modifying AD. Test-Connection returns a simple
# reachable result. Get-CimInstance queries the Domain Controller operating system, and
# Select-Object keeps computer name, OS, build, and boot time readable. A failed check
# raises an error and prevents the rest of onboarding from running."

# ============================================================================
# PHASE 4: CHECK DUPLICATES AND CREATE THE ACTIVE DIRECTORY USER
# Purpose: Import the AD module, generate a username, query for duplicates, and call
# New-ADUser inside the controlled workflow. A secure temporary password is required.
# ============================================================================
function New-EmployeeUsername {
    param([psobject]$Employee)
    return (($Employee.FirstName.Substring(0, 1) + $Employee.LastName) -replace '[^a-zA-Z0-9]', '').ToLowerInvariant()
}
function New-EmployeeAccount {
    param([psobject]$Employee)
    Write-Host '[PHASE 4] Checking duplicates and creating the AD account...'
    Import-Module ActiveDirectory -ErrorAction Stop
    $username = New-EmployeeUsername $Employee
    $existing = Get-ADUser -Filter "SamAccountName -eq '$username' -or UserPrincipalName -eq '$($Employee.Email)'" -Server $script:Config.DomainController -ErrorAction Stop
    if ($existing) { throw "Account already exists for username '$username' or email '$($Employee.Email)'." }
    if ($DryRun) { Write-Host "[PHASE 4] Dry run: $username would be created."; return $username }
    if ($null -eq $TemporaryPassword) { throw 'TemporaryPassword is required for account creation.' }
    $parameters = @{
        Name = "$($Employee.FirstName) $($Employee.LastName)"; GivenName = $Employee.FirstName
        Surname = $Employee.LastName; DisplayName = "$($Employee.FirstName) $($Employee.LastName)"
        SamAccountName = $username; UserPrincipalName = $Employee.Email; Title = $Employee.JobTitle
        Department = $Employee.Department; Description = "Workday Employee ID: $($Employee.EmployeeId)"
        Office = $Employee.Location; OfficePhone = $Employee.PhoneNumber
        AccountPassword = $TemporaryPassword; Enabled = $true; ChangePasswordAtLogon = $true
        Server = $script:Config.DomainController
    }
    New-ADUser @parameters -ErrorAction Stop
    Write-Host "[PHASE 4] Created account $username"
    return $username
}
# Presentation Notes:
# "Phase 4 generates a predictable username and checks both username and email with
# Get-ADUser before New-ADUser runs. A parameter hashtable is splatted into New-ADUser,
# and the secure password is never stored as plain text. Duplicate or permission errors
# are terminating errors handled by the final try/catch."

# ============================================================================
# PHASE 5: DETERMINE AND ASSIGN THE CORRECT ORGANIZATIONAL UNIT
# Purpose: Use an approved Department and EmployeeType key. There is no default OU;
# an unknown mapping stops onboarding for administrator review.
# ============================================================================
function Get-EmployeeOu {
    param([psobject]$Employee)
    $key = "$($Employee.Department)|$($Employee.EmployeeType)"
    if (-not $script:OuMappings.ContainsKey($key)) { throw "No approved OU mapping exists for '$key'. Review required." }
    return $script:OuMappings[$key]
}
function Set-EmployeeOu {
    param([string]$Username, [psobject]$Employee)
    Write-Host '[PHASE 5] Determining and assigning the OU...'
    $ou = Get-EmployeeOu $Employee
    if ($DryRun) { Write-Host "[PHASE 5] Dry run: $Username would move to $ou"; return $ou }
    Get-ADOrganizationalUnit -Identity $ou -Server $script:Config.DomainController -ErrorAction Stop | Out-Null
    $user = Get-ADUser $Username -Server $script:Config.DomainController -ErrorAction Stop
    Move-ADObject -Identity $user.DistinguishedName -TargetPath $ou -Server $script:Config.DomainController -ErrorAction Stop
    Write-Host "[PHASE 5] Assigned $Username to $ou"
    return $ou
}
# Presentation Notes:
# "Phase 5 uses an explicit mapping instead of guessing. Department and employee type
# form a lookup key. If that key is not approved, the script stops. When the OU exists,
# Get-ADOrganizationalUnit verifies it and Move-ADObject places the new user there."

# ============================================================================
# PHASE 6: ASSIGN ROLE-BASED SECURITY GROUPS AND PERMISSIONS
# Purpose: Calculate RBAC groups from approved mappings and process each group with
# foreach. Every successful assignment is retained for the onboarding report.
# ============================================================================
function Get-EmployeeGroups {
    param([psobject]$Employee)
    $groups = [System.Collections.ArrayList]::new()
    foreach ($key in @("Department:$($Employee.Department)", "EmployeeType:$($Employee.EmployeeType)")) {
        if ($script:GroupMappings.ContainsKey($key)) { $script:GroupMappings[$key] | ForEach-Object { [void]$groups.Add($_) } }
    }
    if ($Employee.JobTitle -match '(?i)manager|director|lead') { $script:GroupMappings['JobTitle:Manager'] | ForEach-Object { [void]$groups.Add($_) } }
    return @($groups | Select-Object -Unique)
}
function Add-EmployeeGroups {
    param([string]$Username, [psobject]$Employee)
    Write-Host '[PHASE 6] Assigning role-based groups...'
    foreach ($group in Get-EmployeeGroups $Employee) {
        try {
            if (-not $DryRun) { Add-ADGroupMember -Identity $group -Members $Username -Server $script:Config.DomainController -ErrorAction Stop }
            [void]$script:State.AssignedGroups.Add($group)
            Write-Host "[PHASE 6] Assigned $group"
        }
        catch { throw "Group assignment failed for '$group': $($_.Exception.Message)" }
    }
}
# Presentation Notes:
# "Phase 6 demonstrates role-based access control. Department, employee type, and
# leadership title select approved groups. foreach handles multiple assignments and
# records each success. A failed Add-ADGroupMember call stops the workflow instead of
# silently leaving the user with incomplete permissions."

# ============================================================================
# PHASE 7: PERFORM A RANDOMIZED, READ-ONLY OU ACCESS AUDIT
# Purpose: Sample existing users in the new user's OU. Get-ADUser and
# Get-ADPrincipalGroupMembership read information only. Review findings are never
# automatically remediated.
# ============================================================================
function Invoke-OuAccessAudit {
    param([string]$OrganizationalUnit)
    Write-Host "[PHASE 7] Sampling $AuditSampleSize existing users for a read-only audit..."
    $users = @(Get-ADUser -SearchBase $OrganizationalUnit -Filter * -Properties DisplayName,Enabled,Department,Title,LastLogonDate |
        Where-Object { $_.SamAccountName -ne $script:State.Username } | Get-Random -Count $AuditSampleSize)
    foreach ($user in $users) {
        try {
            $groups = @(Get-ADPrincipalGroupMembership $user -Server $script:Config.DomainController -ErrorAction Stop | Select-Object -ExpandProperty Name)
            $reasons = [System.Collections.ArrayList]::new()
            if (-not $user.Enabled -and $groups.Count -gt 0) { [void]$reasons.Add('Disabled account still has group memberships') }
            if ($user.Department -and $OrganizationalUnit -notmatch [regex]::Escape($user.Department)) { [void]$reasons.Add('Department does not match OU') }
            [pscustomobject]@{
                EmployeeName = $user.DisplayName; Username = $user.SamAccountName
                AccountStatus = if ($user.Enabled) {'Enabled'} else {'Disabled'}
                Department = $user.Department; JobTitle = $user.Title; OU = $OrganizationalUnit
                GroupMemberships = $groups -join '; '; LastLogon = $user.LastLogonDate
                AuditDateTime = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
                ReviewRequired = if ($reasons.Count) {'Yes'} else {'No'}; ReviewReason = $reasons -join '; '
            }
        }
        catch {
            [pscustomobject]@{ EmployeeName = $user.DisplayName; Username = $user.SamAccountName
                AccountStatus = 'Unknown'; Department = $user.Department; JobTitle = $user.Title
                OU = $OrganizationalUnit; GroupMemberships = ''; LastLogon = $user.LastLogonDate
                AuditDateTime = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'; ReviewRequired = 'Yes'
                ReviewReason = "Could not read memberships: $($_.Exception.Message)" }
        }
    }
}
# Presentation Notes:
# "Phase 7 is a governance sample, not a cleanup tool. A random sample limits the work
# during each onboarding event. The audit reads status, department, title, OU, groups,
# and last logon, then marks possible discrepancies for human review. It never disables,
# deletes, or removes access from audited employees."

# ============================================================================
# PHASE 8: EXPORT ONBOARDING, EXCEPTION, AND ACCESS-AUDIT REPORTS
# Purpose: Export CSV files with stable column names for Excel. Exceptions are appended
# by Phase 2; this phase writes successful onboarding and audit results.
# ============================================================================
function Export-Reports {
    param([psobject[]]$AuditRows, [psobject]$Employee, [string]$Username, [string]$OrganizationalUnit)
    Write-Host '[PHASE 8] Exporting CSV reports...'
    New-Item -Path $script:Config.ReportFolder -ItemType Directory -Force | Out-Null
    $onboardingFile = Join-Path $script:Config.ReportFolder 'SuccessfulOnboarding.csv'
    [pscustomobject]@{ EmployeeName = "$($Employee.FirstName) $($Employee.LastName)"; Username = $Username
        AssignedOU = $OrganizationalUnit; AssignedGroups = $script:State.AssignedGroups -join '; '
        DateTime = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'; CompletionStatus = $script:State.Status
        ActionsPerformed = 'Validation; DC check; AD creation; OU assignment; RBAC; OU audit' } |
        Export-Csv $onboardingFile -NoTypeInformation -Append
    $auditFile = Join-Path $script:Config.ReportFolder ("OUAccessAudit_{0}.csv" -f (Get-Date -Format 'yyyy-MM-dd'))
    if ($AuditRows.Count) { $AuditRows | Export-Csv $auditFile -NoTypeInformation }
    else { 'EmployeeName,Username,AccountStatus,Department,JobTitle,OU,GroupMemberships,LastLogon,AuditDateTime,ReviewRequired,ReviewReason' | Set-Content $auditFile }
    Write-Host "[PHASE 8] Reports written to $($script:Config.ReportFolder)"
}
# Presentation Notes:
# "Phase 8 produces practical Excel-compatible reports. SuccessfulOnboarding records
# the new employee, username, OU, groups, actions, time, and status. OUAccessAudit
# contains the sampled account details and review reasons. Phase 2 separately appends
# validation exceptions to the date-based exception file."

# ============================================================================
# PHASE 9: DISPLAY FINAL STATUS AND HANDLE OR LOG ERRORS
# Purpose: try runs the phases in order, catch records serious errors, and finally writes
# a durable log whether the result is successful or failed. Cloud integration is future
# work and is not falsely executed without Graph/Az authentication and Azure resources.
# ============================================================================
try {
    $employee = Import-EmployeeRecord $EmployeeFilePath $EmployeeJson
    Test-EmployeeRecord $employee
    $script:State.EmployeeName = "$($employee.FirstName) $($employee.LastName)"
    Test-DomainController $script:Config.DomainController | Out-Null
    $username = New-EmployeeAccount $employee
    $script:State.Username = $username
    $ou = Set-EmployeeOu $username $employee
    $script:State.AssignedOu = $ou
    Add-EmployeeGroups $username $employee
    $auditRows = if ($DryRun) { @() } else { @(Invoke-OuAccessAudit $ou) }
    if (-not $DryRun) { Export-Reports $auditRows $employee $username $ou }
    $script:State.Status = if ($DryRun) {'Dry Run Completed'} else {'Completed Successfully'}
    Write-Host "[PHASE 9] Onboarding completed: $($script:State.Status)"
}
catch {
    $script:State.Status = 'Failed'
    [void]$script:State.Errors.Add($_.Exception.Message)
    Write-Error "[PHASE 9] Onboarding stopped: $($_.Exception.Message)"
}
finally {
    $script:State.CompletedAt = Get-Date
    New-Item -Path $script:Config.LogFolder -ItemType Directory -Force | Out-Null
    $logFile = Join-Path $script:Config.LogFolder 'Onboarding.log'
    $actions = 'Validation; Domain Controller check; AD account check/creation; OU assignment; RBAC group assignment; Access audit; CSV reporting'
    "[$($script:State.CompletedAt)] Employee: $($script:State.EmployeeName) | Status: $($script:State.Status) | Username: $($script:State.Username) | OU: $($script:State.AssignedOu) | Groups: $($script:State.AssignedGroups -join '; ') | Actions: $actions | Errors: $($script:State.Errors -join '; ')" | Add-Content $logFile
}
# Presentation Notes:
# "Phase 9 ties the workflow together. try runs each phase in order, catch explains
# serious failures, and finally writes a durable log no matter how the run ends. The
# future cloud boundary is intentionally separate: Microsoft Graph and Az automation
# can be added later when authentication, licensing, subscriptions, host pools, and
# images are approved."
