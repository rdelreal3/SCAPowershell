# MSSA Employee Onboarding Presentation Copy
# This script reads one Workday-style employee record, checks it, places the user
# in the correct Active Directory OU, and creates reports for administrators.
#
# ============================================================================
# OVERALL PROJECT EXPLANATION
#
# This tool turns one Workday-style employee record into a controlled
# Active Directory onboarding process. It runs in clear phases so an administrator
# can see what the script is checking, changing, and reporting.
#
# PHASE 1: Receives one employee record from a CSV, JSON value, or parameters.
# PHASE 2: Checks required employee information and records missing fields.
# PHASE 3: Tests LON-DC1 and collects basic operating-system information.
# PHASE 4: Builds a username, checks for duplicates, and creates the account.
# PHASE 5: Uses Department to find and validate the employee's Active Directory OU.
# PHASE 6: Confirms account setup is complete without assigning security groups.
# PHASE 7: Reads a small random sample of existing OU users for review.
# PHASE 8: Saves onboarding, audit, exception, and failure information as CSV files.
# PHASE 9: Displays the result and writes a durable log.
#
# STRUCTURE:
# PowerShell reads commands from top to bottom. The param() block accepts input,
# variables store values, and functions group related work into named reusable steps.
# A hashtable stores labeled settings, an object stores employee information, and
# the pipeline symbol | passes one command's result to the next command.
# Conditions such as if decide whether the script can continue. try, catch, and
# finally handle success, errors, and logging. The main workflow at the bottom
# calls the phases in order and passes information from one phase to the next.
#
# The result is a simple flow:
# Workday record -> validation -> LON-DC1 check -> Department-based OU ->
# duplicate check -> Active Directory account -> read-only audit -> reports and log.
# ============================================================================
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

# ============================================================================
# SCRIPT SETUP / CONFIGURATION
#
# PURPOSE:
# Stores the confirmed lab names and shared-folder locations in one place.
#
# MAIN TOOLS:
# [CmdletBinding()] | param() | [ordered]@{}
#
# WHY IT MATTERS:
# The administrator can review the environment before the script changes anything.
#
# NEXT:
# The script prepares its output folders, then begins Phase 1.
# ============================================================================
# [CmdletBinding()] enables helpful PowerShell command features.
# param() defines the information a person may pass to the script.
# A variable is a named place where PowerShell stores information; variables begin with $.
# [string] means text, [int] means a whole number, and [switch] means an on/off option.
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

# A function is a named block of code that performs one specific job.
# This function checks the existing report share and creates only the Logs subfolder if needed.
function Initialize-OutputFolders {
    # Test-Path checks whether the existing shared report folder can be reached.
    if (-not (Test-Path -LiteralPath $script:Config.ReportFolder)) {
        # throw stops the run with a clear message when the report share is unavailable.
        throw "Report folder is not accessible: $($script:Config.ReportFolder)"
    }
    # The log subfolder may be missing, so New-Item creates only that subfolder.
    if (-not (Test-Path -LiteralPath $script:Config.LogFolder)) {
        # Out-Null hides the folder-creation result because the administrator only needs the status.
        New-Item -Path $script:Config.LogFolder -ItemType Directory -Force | Out-Null
    }
}
# PRESENTATION:
# "The setup section checks that the report share already exists and creates only
# the missing Logs folder. This keeps shared-folder setup predictable and safe."

# Set-StrictMode asks PowerShell to catch misspelled or missing variables early.
# $ErrorActionPreference makes command errors stop the normal workflow.
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
# $script: makes this state available throughout the script.
# [ordered]@{} is a labeled list that keeps its entries in the order shown.
$script:State = [ordered]@{
    EmployeeName = 'Unknown employee'; Username = ''; AssignedOu = ''
    Errors = [System.Collections.ArrayList]::new()
    Status = 'Started'; StartedAt = Get-Date; CompletedAt = $null
}

# ============================================================================
# PHASE 1: RECEIVE EMPLOYEE INFORMATION
#
# PURPOSE:
# Takes one Workday-style CSV or JSON record and loads it into PowerShell.
#
# MAIN TOOLS:
# Test-Path | Import-Csv | Get-Content | ConvertFrom-Json | return
#
# WHY IT MATTERS:
# Workday remains the source of employee information instead of manual retyping.
#
# NEXT:
# Sends exactly one employee record to Phase 2 for validation.
# ============================================================================
function Import-EmployeeRecord {
    param([string]$FilePath, [string]$Json)
    Write-Host '[PHASE 1] Receiving Workday employee information...'
    if ($FilePath) {
        # Test-Path checks whether the supplied file exists before PowerShell opens it.
        # -LiteralPath treats the path exactly as written, and Leaf means a file.
        if (-not (Test-Path -LiteralPath $FilePath -PathType Leaf)) { throw "File not found: $FilePath" }
        # GetExtension reads the ending of the file name so the script can choose its reader.
        $extension = [IO.Path]::GetExtension($FilePath).ToLowerInvariant()
        # @() makes the result an array, even when one record is returned.
        # Import-Csv turns column headings into object properties, so FirstName becomes
        # readable later as $Employee.FirstName.
        if ($extension -eq '.csv') { $records = @(Import-Csv -LiteralPath $FilePath) }
        # Get-Content reads JSON text, and ConvertFrom-Json turns that text into an object.
        elseif ($extension -eq '.json') { $records = @(Get-Content -LiteralPath $FilePath -Raw | ConvertFrom-Json) }
        else { throw "Unsupported file type '$extension'. Use .csv or .json." }
        # .Count tells how many records were loaded; one employee is required per run.
        if ($records.Count -ne 1) { throw "Expected one Workday record, found $($records.Count)." }
        # return sends the one object back to the main workflow.
        return $records[0]
    }
    # JSON passed directly is another supported testing input.
    if ($Json) { return ($Json | ConvertFrom-Json) }
    # [pscustomobject] creates a simple labeled employee object from individual parameters.
    return [pscustomobject]@{
        FirstName = $FirstName; LastName = $LastName; JobTitle = $JobTitle
        Department = $Department; Manager = $Manager; Location = $Location
        Email = $Email; EmployeeType = $EmployeeType; PhoneNumber = $PhoneNumber
        EmployeeId = $EmployeeId
    }
}
# PRESENTATION:
# "Phase 1 reads one employee record from Workday-style input. Import-Csv turns
# column names into information PowerShell can use, and the one-record rule helps
# prevent an accidental bulk onboarding."

# ============================================================================
# PHASE 2: VALIDATE EMPLOYEE INFORMATION
#
# PURPOSE:
# Checks that the employee record contains the information required for onboarding.
#
# MAIN TOOLS:
# [ordered]@{} | foreach | Where-Object | Select-Object | Export-Csv | throw
#
# WHY IT MATTERS:
# Incomplete information must stop before Active Directory is changed.
#
# NEXT:
# Sends complete information to Phase 3, or records the missing fields for HR.
# ============================================================================
function New-ExceptionRecord {
    # Join-Path combines the report folder with the dated exception file name.
    param([string]$EmployeeName, [string]$MissingFields, [string]$ErrorDetails)
    $file = Join-Path $script:Config.ReportFolder ("OnboardingExceptions_{0}.csv" -f (Get-Date -Format 'yyyy-MM-dd'))
    $record = [pscustomobject]@{
        EmployeeName = $EmployeeName; MissingFields = $MissingFields
        DateTime = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'; Status = 'Validation Failed'
        RequiredAction = "HR must update Workday and re-export. Escalate to $($script:Config.HrContact)."
        ErrorDetails = $ErrorDetails
    }
    # if chooses between adding to an existing file and creating a new one.
    if (Test-Path -LiteralPath $file) { $record | Export-Csv $file -NoTypeInformation -Append }
    else { $record | Export-Csv $file -NoTypeInformation }
}
function Test-EmployeeRecord {
    param([Parameter(Mandatory)] [psobject]$Employee)
    Write-Host '[PHASE 2] Validating all required fields...'
    # This ordered hashtable is a labeled list pairing friendly names with object properties.
    $fieldNames = [ordered]@{
        'First Name' = 'FirstName'; 'Last Name' = 'LastName'; 'Job Title' = 'JobTitle'
        Department = 'Department'; Manager = 'Manager'; Location = 'Location'
        Email = 'Email'; 'Phone Number' = 'PhoneNumber'
    }
    $requiredFields = [ordered]@{}
    # foreach repeats the same check for every required field name.
    foreach ($name in $fieldNames.Keys) {
        # PSObject.Properties lets the script safely ask whether a property exists.
        $property = $Employee.PSObject.Properties[$fieldNames[$name]]
        $requiredFields[$name] = if ($property) { $property.Value } else { $null }
    }
    # Step 1: GetEnumerator turns the required-field list into individual items.
    # Step 2: The pipeline symbol | passes those items to Where-Object.
    # Step 3: Where-Object keeps only fields that are blank or missing.
    # Step 4: $_ means the individual field currently being checked.
    # Step 5: Select-Object returns only the name of each missing field.
    # Final result: $missing becomes a list of information HR forgot to provide.
    $missing = @($requiredFields.GetEnumerator() | Where-Object {
        [string]::IsNullOrWhiteSpace([string]$_.Value)
    } | Select-Object -ExpandProperty Key)
    # -gt means "greater than"; a count greater than zero means something is missing.
    if ($missing.Count -gt 0) {
        $name = "$(if ($Employee.PSObject.Properties['FirstName']) {$Employee.FirstName}) $(if ($Employee.PSObject.Properties['LastName']) {$Employee.LastName})".Trim()
        $details = "Missing required fields: $($missing -join ', ')"
        New-ExceptionRecord -EmployeeName $name -MissingFields ($missing -join ', ') -ErrorDetails $details
        throw "$details. HR must correct Workday and re-export. Contact $($script:Config.HrContact)."
    }
    # -not means the condition is false; EmployeeId is optional, so an empty property is added when absent.
    if (-not $Employee.PSObject.Properties['EmployeeId']) { $Employee | Add-Member EmployeeId '' }
    Write-Host '[PHASE 2] Validation succeeded.'
}
# PRESENTATION:
# "Phase 2 is the safety gate. It identifies every missing value, records the exact
# fields in a CSV for review, and stops before incomplete information reaches Active Directory."

# ============================================================================
# PHASE 3: VERIFY THE DOMAIN CONTROLLER AND COLLECT CIM INFORMATION
#
# PURPOSE:
# Checks that LON-DC1 is reachable and asks it for basic operating-system information.
#
# MAIN TOOLS:
# Test-Connection | Get-CimInstance | Select-Object
#
# WHY IT MATTERS:
# Infrastructure is checked before the script attempts an Active Directory change.
#
# NEXT:
# Sends the successful infrastructure check to Phase 5 for OU validation.
# ============================================================================
function Test-DomainController {
    param([string]$ComputerName)
    Write-Host "[PHASE 3] Checking $ComputerName with Test-Connection..."
    # Test-Connection checks whether LON-DC1 can be reached across the network.
    # -Quiet returns only True or False instead of full ping details.
    # -not reverses True to False; -ErrorAction Stop turns a command error into a catchable failure.
    if (-not (Test-Connection -ComputerName $ComputerName -Count 2 -Quiet)) { throw "Cannot reach $ComputerName. Onboarding stopped." }
    # Get-CimInstance asks the server for management information about its operating system.
    # Win32_OperatingSystem is the Windows information category being requested.
    # The pipeline symbol | sends the result into Select-Object.
    $info = Get-CimInstance Win32_OperatingSystem -ComputerName $ComputerName -ErrorAction Stop |
        # Select-Object shows only the pieces of information we care about.
        Select-Object @{Name='ComputerName';Expression={$_.CSName}}, Caption, Version, BuildNumber, LastBootUpTime
    Write-Host "[PHASE 3] $($info.ComputerName): $($info.Caption), build $($info.BuildNumber)"
    return $info
}
# PRESENTATION:
# "Phase 3 checks the Domain Controller before account work begins. It confirms the
# server responds, then collects a small, readable summary of its operating system."

# ============================================================================
# PHASE 4: CHECK DUPLICATES AND CREATE THE ACTIVE DIRECTORY USER
#
# PURPOSE:
# Builds a username, checks for an existing account, and creates the new account.
#
# MAIN TOOLS:
# Get-ADUser | New-ADUser | ConvertTo-SecureString | RandomNumberGenerator
#
# WHY IT MATTERS:
# Duplicate accounts are blocked, and each real account receives a new initial password.
#
# NEXT:
# Sends the created account name to Phase 6 for completion status.
# ============================================================================
function New-EmployeeUsername {
    param([psobject]$Employee)
    # Substring(0,1) takes the first letter of the first name.
    # The + operator combines that letter with the last name.
    # -replace removes characters that are not letters or numbers.
    # ToLowerInvariant makes the username lowercase. Chris McGhee becomes cmcghee.
    return (($Employee.FirstName.Substring(0, 1) + $Employee.LastName) -replace '[^a-zA-Z0-9]', '').ToLowerInvariant()
}
function New-InitialPassword {
    # These four strings hold uppercase letters, lowercase letters, numbers, and special characters.
    $characterSets = @(
        'ABCDEFGHIJKLMNOPQRSTUVWXYZ',
        'abcdefghijklmnopqrstuvwxyz',
        '0123456789',
        '!@#$%^&*()-_=+[]{}'
    )
    # This list holds password characters before they become one string.
    $characters = [System.Collections.Generic.List[char]]::new()
    # RandomNumberGenerator is a built-in .NET tool for strong random values.
    $random = [System.Security.Cryptography.RandomNumberGenerator]::Create()
    try {
        # foreach repeats the same action once for each character category.
        foreach ($characterSet in $characterSets) {
            # byte[] is a tiny container for a random number generated by Windows.
            $bytes = New-Object byte[] 1
            $random.GetBytes($bytes)
            # % keeps the random number inside the valid range of character positions.
            [void]$characters.Add($characterSet[$bytes[0] % $characterSet.Length])
        }
        $allCharacters = ($characterSets -join '')
        # while repeats until the password reaches 16 characters.
        while ($characters.Count -lt 16) {
            $bytes = New-Object byte[] 1
            $random.GetBytes($bytes)
            [void]$characters.Add($allCharacters[$bytes[0] % $allCharacters.Length])
        }
        # This for loop shuffles the characters into a random order.
        for ($index = $characters.Count - 1; $index -gt 0; $index--) {
            $bytes = New-Object byte[] 1
            $random.GetBytes($bytes)
            $swapIndex = $bytes[0] % ($index + 1)
            $temporary = $characters[$index]
            $characters[$index] = $characters[$swapIndex]
            $characters[$swapIndex] = $temporary
        }
        # -join combines the individual characters into one complete password.
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
    # Get-ADUser checks whether the username or email already exists.
    $existing = Get-ADUser -Filter "SamAccountName -eq '$username' -or UserPrincipalName -eq '$($Employee.Email)'" -Server $script:Config.DomainController -ErrorAction Stop
    # if stops when the duplicate check returns an existing account.
    if ($existing) { throw "Account already exists for username '$username' or email '$($Employee.Email)'." }
    # DryRun performs the safety checks without creating the Active Directory account.
    # if DryRun is True, the function returns before password generation or account creation.
    if ($DryRun) { Write-Host "[PHASE 4] Dry run: $username would be created."; return $username }
    # The password temporarily exists as text so it can be displayed for this lab.
    # ConvertTo-SecureString creates the password format New-ADUser expects.
    $initialPassword = New-InitialPassword
    $secureInitialPassword = ConvertTo-SecureString -String $initialPassword -AsPlainText -Force
    # This hashtable is a labeled list of information sent to New-ADUser.
    $parameters = @{
        Name = "$($Employee.FirstName) $($Employee.LastName)"; GivenName = $Employee.FirstName
        Surname = $Employee.LastName; DisplayName = "$($Employee.FirstName) $($Employee.LastName)"
        SamAccountName = $username; UserPrincipalName = $Employee.Email; Title = $Employee.JobTitle
        Department = $Employee.Department; Description = "Workday Employee ID: $($Employee.EmployeeId)"
        Office = $Employee.Location; OfficePhone = $Employee.PhoneNumber
        # Enabled makes the account active; ChangePasswordAtLogon requires a new password at first sign-in.
        AccountPassword = $secureInitialPassword; Enabled = $true; ChangePasswordAtLogon = $true
        # Server selects the Domain Controller; Path selects the validated OU.
        Server = $script:Config.DomainController; Path = $OrganizationalUnit
    }
    # @parameters is splatting: PowerShell sends all saved settings to New-ADUser.
    New-ADUser @parameters -ErrorAction Stop
    Write-Host "[PHASE 4] Created account $username"
    Write-Host "INITIAL PASSWORD FOR ${username}: $initialPassword"
    return $username
}
# PRESENTATION:
# "Phase 4 creates a predictable username and checks for duplicates first. On a real
# run, a strong initial password is sent to Active Directory and must be changed at first sign-in."

# ============================================================================
# PHASE 5: DETERMINE AND VALIDATE THE CORRECT ORGANIZATIONAL UNIT
#
# PURPOSE:
# Uses only Department to find the existing Active Directory OU for the employee.
#
# MAIN TOOLS:
# Get-ADOrganizationalUnit | Select-Object | -notcontains
#
# WHY IT MATTERS:
# The correct OU must be known before Phase 4 creates the account.
#
# NEXT:
# Sends the validated OU address to Phase 4 for direct account creation.
# ============================================================================
function Get-EmployeeOu {
    param([psobject]$Employee)
    # This array is the list of department names the script is allowed to use.
    $approvedDepartments = @('Development', 'IT', 'HR', 'Marketing', 'Research', 'Sales')
    $department = [string]$Employee.Department
    # -notcontains checks whether the employee department is missing from the approved list.
    # if the controlled department is not approved, the script stops instead of guessing.
    if ($approvedDepartments -notcontains $department) {
        throw "Workday Department '$department' is not an approved department OU name. Onboarding stopped."
    }
    # Get-ADOrganizationalUnit searches Active Directory for an OU matching the department.
    # Select-Object keeps the OU name and its DistinguishedName, the full AD address.
    # For example: OU=HR,DC=Adatum,DC=com.
    $organizationalUnits = @(Get-ADOrganizationalUnit -Filter "Name -eq '$department'" -Server $script:Config.DomainController -ErrorAction Stop |
        Select-Object Name, DistinguishedName)
    # Zero matches means stop because no destination was found.
    # -eq means "equals"; zero matching OUs is a stop condition.
    if ($organizationalUnits.Count -eq 0) {
        throw "No Organizational Unit named '$department' was found in Active Directory. Onboarding stopped."
    }
    # More than one match means stop because the destination is not unambiguous.
    # A count greater than one is also a stop condition because the choice is ambiguous.
    if ($organizationalUnits.Count -gt 1) {
        throw "Multiple Organizational Units named '$department' were found in Active Directory. Onboarding stopped until the OU is unambiguous."
    }
    # Exactly one match continues and returns the OU's full Active Directory address.
    return $organizationalUnits[0].DistinguishedName
}
function Set-EmployeeOu {
    param([psobject]$Employee)
    Write-Host '[PHASE 5] Determining and validating the target OU...'
    # Phase 5 checks the OU before Phase 4 creates the account.
    $ou = Get-EmployeeOu $Employee
    Get-ADOrganizationalUnit -Identity $ou -Server $script:Config.DomainController -ErrorAction Stop | Out-Null
    # if DryRun is True, show the planned OU without changing Active Directory.
    if ($DryRun) { Write-Host "[PHASE 5] Dry run: a user would be created in $ou" }
    else { Write-Host "[PHASE 5] Validated target OU $ou" }
    return $ou
}
# PRESENTATION:
# "Phase 5 is the safety step for placement. Department is matched to one existing OU;
# zero or multiple matches stop the run, so the script never guesses where a user belongs."

# ============================================================================
# PHASE 6: COMPLETE ACCOUNT SETUP
#
# PURPOSE:
# Marks the account setup portion as complete after the account name is known.
#
# MAIN TOOLS:
# Write-Host
#
# WHY IT MATTERS:
# This version keeps the nine-phase presentation structure simple and clear.
# It does not automatically assign security groups.
#
# NEXT:
# Sends the account to Phase 7 for a read-only review of existing users.
# ============================================================================
# PRESENTATION:
# "The account creation portion is complete. This version does not automatically
# assign security groups, so the process stays focused on the employee account and OU."

# ============================================================================
# PHASE 7: PERFORM A RANDOMIZED, READ-ONLY OU AUDIT
#
# PURPOSE:
# Reviews a small random sample of existing users in the new employee's OU.
#
# MAIN TOOLS:
# Get-ADUser | Where-Object | Math.Min | Get-Random | Get-ADPrincipalGroupMembership
#
# WHY IT MATTERS:
# Administrators receive useful review information without changing existing accounts.
#
# NEXT:
# Sends audit rows to Phase 8 for CSV reporting.
# ============================================================================
# THIS AUDIT IS READ-ONLY.
# It does not disable users, delete accounts, change group memberships, or change permissions.
function Invoke-OuAccessAudit {
    param([string]$OrganizationalUnit, [string]$ExpectedDepartment)
    # Get-ADUser -SearchBase reads users from only the selected OU.
    # Where-Object excludes the new employee from the existing-user sample.
    $eligibleUsers = @(Get-ADUser -SearchBase $OrganizationalUnit -Filter * -Properties DisplayName,Enabled,Department,Title,LastLogonDate |
        Where-Object { $_.SamAccountName -ne $script:State.Username })
    # Math.Min prevents the script from asking for more users than actually exist.
    $sampleCount = [Math]::Min($AuditSampleSize, $eligibleUsers.Count)
    # -eq means "equals"; no eligible users produces an empty audit instead of an error.
    if ($sampleCount -eq 0) {
        Write-Host "[PHASE 7] No eligible existing users were found in $OrganizationalUnit; audit returned no rows."
        return @()
    }
    Write-Host "[PHASE 7] Sampling $sampleCount of $($eligibleUsers.Count) eligible existing users for a read-only audit..."
    # Get-Random selects a small random group of existing users for review.
    $users = @($eligibleUsers | Get-Random -Count $sampleCount)
    foreach ($user in $users) {
        try {
            # This reads the selected user's group memberships; it does not change them.
            $groups = @(Get-ADPrincipalGroupMembership $user -Server $script:Config.DomainController -ErrorAction Stop | Select-Object -ExpandProperty Name)
            $reasons = [System.Collections.ArrayList]::new()
            # -and requires both conditions; -ne means "does not equal" for the department check.
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
# PRESENTATION:
# "Phase 7 takes a small random sample for review. It only reads account and membership
# information and flags possible issues for an administrator; it never changes user access."

# ============================================================================
# PHASE 8: EXPORT ONBOARDING, EXCEPTION, AND AUDIT REPORTS
#
# PURPOSE:
# Saves the onboarding result and read-only audit results as easy-to-open CSV files.
#
# MAIN TOOLS:
# Join-Path | [pscustomobject] | Export-Csv | -Append | -NoTypeInformation
#
# WHY IT MATTERS:
# Administrators can review the results in Excel without searching through console output.
#
# NEXT:
# Sends the completed or failed result to Phase 9 for status and logging.
# ============================================================================
function Export-Reports {
    param([psobject[]]$AuditRows, [psobject]$Employee, [string]$Username, [string]$OrganizationalUnit)
    Write-Host '[PHASE 8] Exporting CSV reports...'
    # Join-Path combines the shared folder location with a file name.
    $onboardingFile = Join-Path $script:Config.ReportFolder 'SuccessfulOnboarding.csv'
    # [pscustomobject] creates one structured row with labeled columns.
    [pscustomobject]@{ EmployeeName = "$($Employee.FirstName) $($Employee.LastName)"; Username = $Username
        AssignedOU = $OrganizationalUnit
        DateTime = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'; CompletionStatus = $script:State.Status
        ActionsPerformed = 'Validation; DC check; AD creation; OU assignment; OU audit' } |
        # Export-Csv saves PowerShell information in a CSV that can be opened in Excel.
        # -Append adds a row without deleting older records; -NoTypeInformation removes extra metadata.
        Export-Csv $onboardingFile -NoTypeInformation -Append
    $auditFile = Join-Path $script:Config.ReportFolder ("OUAccessAudit_{0}.csv" -f (Get-Date -Format 'yyyy-MM-dd'))
    if ($AuditRows.Count) {
        # The daily audit file also appends new rows so previous results are preserved.
        if (Test-Path -LiteralPath $auditFile) { $AuditRows | Export-Csv $auditFile -NoTypeInformation -Append }
        else { $AuditRows | Export-Csv $auditFile -NoTypeInformation }
    }
    elseif (-not (Test-Path -LiteralPath $auditFile)) {
        'EmployeeName,Username,AccountStatus,Department,JobTitle,OU,GroupMemberships,LastLogon,AuditDateTime,ReviewRequired,ReviewReason' | Set-Content $auditFile
    }
    Write-Host "[PHASE 8] Reports written to $($script:Config.ReportFolder)"
}
# PRESENTATION:
# "Phase 8 turns the results into CSV files. SuccessfulOnboarding records the new
# account, OU, and status; the daily audit and exception files preserve review history."

# ============================================================================
# PHASE 9: DISPLAY FINAL STATUS AND HANDLE ERRORS
#
# PURPOSE:
# Gives a clear result and records what happened in a durable report and log.
#
# MAIN TOOLS:
# try | catch | finally | Export-FailureReport | Add-Content
#
# WHY IT MATTERS:
# Failures are recorded for review instead of disappearing from the console.
#
# NEXT:
# The onboarding attempt is complete and can be reviewed by an administrator.
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
    # try is the normal path the script attempts to complete.
    # Initialize-OutputFolders checks the shared report folder and prepares logging.
    Initialize-OutputFolders
    # Phase 1 reads the file and returns the employee information.
    $employee = Import-EmployeeRecord $EmployeeFilePath $EmployeeJson
    # Phase 2 checks the employee information for missing values.
    Test-EmployeeRecord $employee
    $script:State.EmployeeName = "$($employee.FirstName) $($employee.LastName)"
    # Phase 3 checks LON-DC1 and collects operating-system information.
    Test-DomainController $script:Config.DomainController | Out-Null
    # Phase 5 finds and returns the correct Active Directory OU.
    $ou = Set-EmployeeOu $employee
    $script:State.AssignedOu = $ou
    # Phase 4 uses the employee and validated OU to create the account, or describe the dry run.
    $username = New-EmployeeAccount $employee $ou
    $script:State.Username = $username
    # Phase 6 reports that account setup is complete; no security groups are assigned.
    Write-Host "[PHASE 6] Account setup complete for $username; no security-group assignment is configured."
    $script:State.Status = if ($DryRun) {'Dry Run Completed'} else {'Completed Successfully'}
    $auditRows = if ($DryRun) { @() } else { @(Invoke-OuAccessAudit $ou $employee.Department) }
    if (-not $DryRun) { Export-Reports $auditRows $employee $username $ou }
    Write-Host "[PHASE 9] Onboarding completed: $($script:State.Status)"
}
catch {
    # catch runs when something fails and records the error instead of failing silently.
    $script:State.Status = 'Failed'
    [void]$script:State.Errors.Add($_.Exception.Message)
    Export-FailureReport -ErrorDetails $_.Exception.Message
    Write-Error "[PHASE 9] Onboarding stopped: $($_.Exception.Message)"
}
finally {
    # finally runs whether the script succeeds or fails, so every attempt can be logged.
    $script:State.CompletedAt = Get-Date
    $logFile = Join-Path $script:Config.LogFolder 'Onboarding.log'
    $actions = 'Validation; Domain Controller check; AD account check/creation; OU assignment; Access audit; CSV reporting'
    "[$($script:State.CompletedAt)] Employee: $($script:State.EmployeeName) | Status: $($script:State.Status) | Username: $($script:State.Username) | OU: $($script:State.AssignedOu) | Actions: $actions | Errors: $($script:State.Errors -join '; ')" | Add-Content $logFile
}
# PRESENTATION:
# "Phase 9 gives the administrator a clear success or failure result. try handles the
# normal path, catch records problems, and finally writes a log whether the run succeeds or fails."