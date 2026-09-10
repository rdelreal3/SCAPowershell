# MSSA Employee Onboarding, Identity Access, and Access Audit Tool
# Workflow: Workday Export -> Power Automate -> PowerShell -> Active Directory -> Reports
# Phase 1: Receive employee information.
# Phase 2: Validate employee information and log exceptions.
# Phase 3: Verify the Domain Controller and collect CIM information.
# Phase 4: Check for duplicates and create the AD user.
# Phase 5: Determine and assign the correct OU.
# Phase 6: Complete account setup.
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
    [int]$AuditSampleSize = 0,
    [switch]$DryRun
)

# Confirmed lab environment values are kept here so administrators can review one section.
# ADATUM\OnboardingAutomation must be delegated only the permissions required to query
# users, create users in approved OUs, read OUs, and read user/group information for audits.
# Do not make it a Domain Admin or store its password here.
$script:Config = [ordered]@{
    DomainController = 'LON-DC1'
    DomainName = 'Adatum.com'
    ClientComputer = 'LON-CL1'
    ServerComputer = 'LON-SVR1'
    WorkdayFolder = '\\LON-SVR1\WorkdayExports'
    ReportFolder = '\\LON-SVR1\ITReports'
    LogFolder = '\\LON-SVR1\ITReports\Logs'
    HrContact = 'Jordan Mitchell | HR Onboarding Director | jordan.mitchell@adatum.com | (555) 014-7284'
    DefaultAuditSampleSize = 3
}
if ($AuditSampleSize -le 0) { $AuditSampleSize = $script:Config.DefaultAuditSampleSize }

function Initialize-OutputFolders {
    if (-not (Test-Path -LiteralPath $script:Config.ReportFolder)) {
        throw "Report folder is not accessible: $($script:Config.ReportFolder)"
    }
    if (-not (Test-Path -LiteralPath $script:Config.LogFolder)) {
        New-Item -Path $script:Config.LogFolder -ItemType Directory -Force | Out-Null
    }
}

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$script:State = [ordered]@{
    EmployeeName = 'Unknown employee'; Username = ''; AssignedOu = ''
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
        Email = 'Email'; 'Phone Number' = 'PhoneNumber'
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
# New-ADUser inside the controlled workflow with a generated initial password.
# ============================================================================
function New-EmployeeUsername {
    param([psobject]$Employee)
    return (($Employee.FirstName.Substring(0, 1) + $Employee.LastName) -replace '[^a-zA-Z0-9]', '').ToLowerInvariant()
}
function New-InitialPassword {
    $characterSets = @(
        'ABCDEFGHIJKLMNOPQRSTUVWXYZ',
        'abcdefghijklmnopqrstuvwxyz',
        '0123456789',
        '!@#$%^&*()-_=+[]{}'
    )
    $characters = [System.Collections.Generic.List[char]]::new()
    $random = [System.Security.Cryptography.RandomNumberGenerator]::Create()
    try {
        foreach ($characterSet in $characterSets) {
            $bytes = New-Object byte[] 1
            $random.GetBytes($bytes)
            [void]$characters.Add($characterSet[$bytes[0] % $characterSet.Length])
        }
        $allCharacters = ($characterSets -join '')
        while ($characters.Count -lt 16) {
            $bytes = New-Object byte[] 1
            $random.GetBytes($bytes)
            [void]$characters.Add($allCharacters[$bytes[0] % $allCharacters.Length])
        }
        for ($index = $characters.Count - 1; $index -gt 0; $index--) {
            $bytes = New-Object byte[] 1
            $random.GetBytes($bytes)
            $swapIndex = $bytes[0] % ($index + 1)
            $temporary = $characters[$index]
            $characters[$index] = $characters[$swapIndex]
            $characters[$swapIndex] = $temporary
        }
        return -join $characters
    }
    finally {
        $random.Dispose()
    }
}
function New-EmployeeAccount {
    param([psobject]$Employee, [string]$OrganizationalUnit)
    Write-Host '[PHASE 4] Checking duplicates and creating the AD account...'
    Import-Module ActiveDirectory -ErrorAction Stop
    $username = New-EmployeeUsername $Employee
    $existing = Get-ADUser -Filter "SamAccountName -eq '$username' -or UserPrincipalName -eq '$($Employee.Email)'" -Server $script:Config.DomainController -ErrorAction Stop
    if ($existing) { throw "Account already exists for username '$username' or email '$($Employee.Email)'." }
    if ($DryRun) { Write-Host "[PHASE 4] Dry run: $username would be created."; return $username }
    $initialPassword = New-InitialPassword
    $secureInitialPassword = ConvertTo-SecureString -String $initialPassword -AsPlainText -Force
    $parameters = @{
        Name = "$($Employee.FirstName) $($Employee.LastName)"; GivenName = $Employee.FirstName
        Surname = $Employee.LastName; DisplayName = "$($Employee.FirstName) $($Employee.LastName)"
        SamAccountName = $username; UserPrincipalName = $Employee.Email; Title = $Employee.JobTitle
        Department = $Employee.Department; Description = "Workday Employee ID: $($Employee.EmployeeId)"
        Office = $Employee.Location; OfficePhone = $Employee.PhoneNumber
        AccountPassword = $secureInitialPassword; Enabled = $true; ChangePasswordAtLogon = $true
        Server = $script:Config.DomainController; Path = $OrganizationalUnit
    }
    New-ADUser @parameters -ErrorAction Stop
    Write-Host "[PHASE 4] Created account $username"
    Write-Host "INITIAL PASSWORD FOR ${username}: $initialPassword"
    return $username
}
# Presentation Notes:
# "Phase 4 generates a predictable username and checks both username and email with
# Get-ADUser before New-ADUser runs. A generated initial password is converted to a
# SecureString and displayed once only after successful account creation."

# ============================================================================
# PHASE 5: DETERMINE AND ASSIGN THE CORRECT ORGANIZATIONAL UNIT
# Purpose: Use the controlled Workday Department value to find the matching OU by name.
# There is no default OU; a missing or ambiguous match stops onboarding.
# ============================================================================
function Get-EmployeeOu {
    param([psobject]$Employee)
    $approvedDepartments = @('Development', 'IT', 'HR', 'Marketing', 'Research', 'Sales')
    $department = [string]$Employee.Department
    if ($approvedDepartments -notcontains $department) {
        throw "Workday Department '$department' is not an approved department OU name. Onboarding stopped."
    }
    $organizationalUnits = @(Get-ADOrganizationalUnit -Filter "Name -eq '$department'" -Server $script:Config.DomainController -ErrorAction Stop |
        Select-Object Name, DistinguishedName)
    if ($organizationalUnits.Count -eq 0) {
        throw "No Organizational Unit named '$department' was found in Active Directory. Onboarding stopped."
    }
    if ($organizationalUnits.Count -gt 1) {
        throw "Multiple Organizational Units named '$department' were found in Active Directory. Onboarding stopped until the OU is unambiguous."
    }
    return $organizationalUnits[0].DistinguishedName
}
function Set-EmployeeOu {
    param([psobject]$Employee)
    Write-Host '[PHASE 5] Determining and validating the target OU...'
    $ou = Get-EmployeeOu $Employee
    Get-ADOrganizationalUnit -Identity $ou -Server $script:Config.DomainController -ErrorAction Stop | Out-Null
    if ($DryRun) { Write-Host "[PHASE 5] Dry run: a user would be created in $ou" }
    else { Write-Host "[PHASE 5] Validated target OU $ou" }
    return $ou
}
# Presentation Notes:
# "Phase 5 uses the controlled Workday Department value as the exact OU name. If no
# matching OU exists, or more than one match exists, the script stops. The validated
# DistinguishedName is passed to New-ADUser so no post-creation move is required."

# ============================================================================
# PHASE 6: COMPLETE ACCOUNT SETUP
# Purpose: Keep the nine-phase project structure without assigning security groups.
# ============================================================================

# ============================================================================
# PHASE 7: PERFORM A RANDOMIZED, READ-ONLY OU ACCESS AUDIT
# Purpose: Sample existing users in the new user's OU. Get-ADUser and
# Get-ADPrincipalGroupMembership read information only. Review findings are never
# automatically remediated.
# ============================================================================
function Invoke-OuAccessAudit {
    param([string]$OrganizationalUnit, [string]$ExpectedDepartment)
    $eligibleUsers = @(Get-ADUser -SearchBase $OrganizationalUnit -Filter * -Properties DisplayName,Enabled,Department,Title,LastLogonDate |
        Where-Object { $_.SamAccountName -ne $script:State.Username })
    $sampleCount = [Math]::Min($AuditSampleSize, $eligibleUsers.Count)
    if ($sampleCount -eq 0) {
        Write-Host "[PHASE 7] No eligible existing users were found in $OrganizationalUnit; audit returned no rows."
        return @()
    }
    Write-Host "[PHASE 7] Sampling $sampleCount of $($eligibleUsers.Count) eligible existing users for a read-only audit..."
    $users = @($eligibleUsers | Get-Random -Count $sampleCount)
    foreach ($user in $users) {
        try {
            $groups = @(Get-ADPrincipalGroupMembership $user -Server $script:Config.DomainController -ErrorAction Stop | Select-Object -ExpandProperty Name)
            $reasons = [System.Collections.ArrayList]::new()
            if (-not $user.Enabled -and $groups.Count -gt 0) { [void]$reasons.Add('Disabled account still has group memberships') }
            if ($user.Department -and $user.Department -ne $ExpectedDepartment) { [void]$reasons.Add('Department does not match OU') }
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
    $onboardingFile = Join-Path $script:Config.ReportFolder 'SuccessfulOnboarding.csv'
    [pscustomobject]@{ EmployeeName = "$($Employee.FirstName) $($Employee.LastName)"; Username = $Username
        AssignedOU = $OrganizationalUnit
        DateTime = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'; CompletionStatus = $script:State.Status
        ActionsPerformed = 'Validation; DC check; AD creation; OU assignment; OU audit' } |
        Export-Csv $onboardingFile -NoTypeInformation -Append
    $auditFile = Join-Path $script:Config.ReportFolder ("OUAccessAudit_{0}.csv" -f (Get-Date -Format 'yyyy-MM-dd'))
    if ($AuditRows.Count) {
        if (Test-Path -LiteralPath $auditFile) { $AuditRows | Export-Csv $auditFile -NoTypeInformation -Append }
        else { $AuditRows | Export-Csv $auditFile -NoTypeInformation }
    }
    elseif (-not (Test-Path -LiteralPath $auditFile)) {
        'EmployeeName,Username,AccountStatus,Department,JobTitle,OU,GroupMemberships,LastLogon,AuditDateTime,ReviewRequired,ReviewReason' | Set-Content $auditFile
    }
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
# a failure CSV and durable log whether the result is successful or failed. Cloud integration is future
# work and is not falsely executed without Graph/Az authentication and Azure resources.
# ============================================================================
function Export-FailureReport {
    param([string]$ErrorDetails)
    $failureFile = Join-Path $script:Config.ReportFolder ("OnboardingFailures_{0}.csv" -f (Get-Date -Format 'yyyy-MM-dd'))
    $record = [pscustomobject]@{
        EmployeeName = $script:State.EmployeeName
        Username = $script:State.Username
        AssignedOU = $script:State.AssignedOu
        Status = 'Failed'
        ErrorDetails = $ErrorDetails
        DateTime = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
        RequiredAction = 'Review the failure, correct the source or environment issue, and re-run onboarding.'
    }
    if (Test-Path -LiteralPath $failureFile) { $record | Export-Csv $failureFile -NoTypeInformation -Append }
    else { $record | Export-Csv $failureFile -NoTypeInformation }
}
try {
    Initialize-OutputFolders
    $employee = Import-EmployeeRecord $EmployeeFilePath $EmployeeJson
    Test-EmployeeRecord $employee
    $script:State.EmployeeName = "$($employee.FirstName) $($employee.LastName)"
    Test-DomainController $script:Config.DomainController | Out-Null
    $ou = Set-EmployeeOu $employee
    $script:State.AssignedOu = $ou
    $username = New-EmployeeAccount $employee $ou
    $script:State.Username = $username
    Write-Host "[PHASE 6] Account setup complete for $username; no security-group assignment is configured."
    $script:State.Status = if ($DryRun) {'Dry Run Completed'} else {'Completed Successfully'}
    $auditRows = if ($DryRun) { @() } else { @(Invoke-OuAccessAudit $ou $employee.Department) }
    if (-not $DryRun) { Export-Reports $auditRows $employee $username $ou }
    Write-Host "[PHASE 9] Onboarding completed: $($script:State.Status)"
}
catch {
    $script:State.Status = 'Failed'
    [void]$script:State.Errors.Add($_.Exception.Message)
    Export-FailureReport -ErrorDetails $_.Exception.Message
    Write-Error "[PHASE 9] Onboarding stopped: $($_.Exception.Message)"
}
finally {
    $script:State.CompletedAt = Get-Date
    $logFile = Join-Path $script:Config.LogFolder 'Onboarding.log'
    $actions = 'Validation; Domain Controller check; AD account check/creation; OU assignment; Access audit; CSV reporting'
    "[$($script:State.CompletedAt)] Employee: $($script:State.EmployeeName) | Status: $($script:State.Status) | Username: $($script:State.Username) | OU: $($script:State.AssignedOu) | Actions: $actions | Errors: $($script:State.Errors -join '; ')" | Add-Content $logFile
}
# Presentation Notes:
# "Phase 9 ties the workflow together. try runs each phase in order, catch explains
# serious failures and exports a failure CSV, and finally writes a durable log no matter how the run ends. The
# future cloud boundary is intentionally separate: Microsoft Graph and Az automation
# can be added later when authentication, licensing, subscriptions, host pools, and
# images are approved."
