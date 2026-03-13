<#
.SYNOPSIS
    Exportiert alle Entra ID Gruppen mit Properties, Mitgliedschaften, PIM-Einstellungen
    und Access-Review-Informationen in mehrere CSV-Dateien.

.DESCRIPTION
    Das Script verbindet sich mit Microsoft Graph (Least Privilege / Zero Trust) und erstellt:

    1. Gruppen_Uebersicht.csv  - Alle Gruppen mit Properties
    2. Gruppen_Mitglieder.csv  - Mitgliedschaften (Members + Owners) pro Gruppe
    3. PIM_Einstellungen.csv   - Detaillierte PIM-Settings fuer aktivierte Gruppen
    4. PIM_Zuweisungen.csv     - Aktive und berechtigte PIM-Zuweisungen

    Exportierte Gruppen-Properties:
      - Group Name, Description, Group Type, Membership Type
      - Rollenzuweisung moeglich (isAssignableToRole)
      - PIM aktiviert (Ja/Nein)

    PIM-Details (nur fuer PIM-aktivierte Gruppen):
      - Activation Settings (Max Duration, MFA, Justification, Ticketing, Approval, Approvers)
      - Assignment Settings (Permanent Eligible/Active, Expiration, MFA, Justification)
      - Active Assignments, Eligible Assignments
      - Access Review Konfiguration

.PARAMETER OutputFolder
    Ordner fuer die Export-Dateien. Standard: .\EntraID_Export_<Datum>

.PARAMETER TenantId
    Optionale Tenant-ID fuer Multi-Tenant-Umgebungen.

.EXAMPLE
    .\Export-EntraIDGroups.ps1
    .\Export-EntraIDGroups.ps1 -TenantId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
    .\Export-EntraIDGroups.ps1 -OutputFolder "C:\Reports\EntraExport"

.NOTES
    Berechtigungen (Least Privilege / Zero Trust):
      - Group.Read.All                                    (Gruppen + Members + Owners lesen)
      - PrivilegedEligibilitySchedule.Read.AzureADGroup   (PIM Eligible Schedules)
      - PrivilegedAssignmentSchedule.Read.AzureADGroup    (PIM Assignment Schedules)
      - RoleManagementPolicy.Read.AzureADGroup            (PIM Policy Rules)
      - AccessReview.Read.All                             (Access Reviews)

    Voraussetzung: Microsoft.Graph PowerShell SDK (Install-Module Microsoft.Graph)
#>

[CmdletBinding()]
param(
    [string]$OutputFolder = ".\EntraID_Export_$(Get-Date -Format 'yyyyMMdd_HHmmss')",
    [string]$TenantId = ""
)

#region ── Hilfsfunktionen ───────────────────────────────────────────────────

function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $color = switch ($Level) {
        "WARN"  { "Yellow" }
        "ERROR" { "Red" }
        "OK"    { "Green" }
        default { "Cyan" }
    }
    Write-Host "[$timestamp] [$Level] $Message" -ForegroundColor $color
}

function Get-GraphPagedResults {
    param([string]$Uri)
    $results = @()
    $nextLink = $Uri
    do {
        try {
            $response = Invoke-MgGraphRequest -Uri $nextLink -Method GET -ErrorAction Stop
            if ($response.value) { $results += $response.value }
            $nextLink = $response.'@odata.nextLink'
        }
        catch {
            Write-Log "Graph-Fehler bei URI: $nextLink - $_" -Level "ERROR"
            break
        }
    } while ($nextLink)
    return $results
}

function Get-GroupTypeName {
    param($Group)
    $gt          = $Group.groupTypes
    $mailEnabled = $Group.mailEnabled
    $secEnabled  = $Group.securityEnabled

    if ($gt -contains "Unified")                       { return "Microsoft 365" }
    elseif ($secEnabled -and $mailEnabled)             { return "Mail-enabled Security" }
    elseif ($secEnabled -and -not $mailEnabled)        { return "Security" }
    elseif ($mailEnabled -and -not $secEnabled)        { return "Distribution" }
    else                                               { return "Unbekannt" }
}

function Get-MembershipTypeName {
    param($Group)
    if ($Group.groupTypes -contains "DynamicMembership") { return "Dynamisch" }
    else                                                 { return "Zugewiesen" }
}

function ConvertFrom-IsoDuration {
    param([string]$Duration)
    if (-not $Duration) { return "" }
    if ($Duration -match '^P(?:(\d+)Y)?(?:(\d+)M)?(?:(\d+)D)?(?:T(?:(\d+)H)?(?:(\d+)M)?(?:(\d+)S)?)?$') {
        $parts = @()
        if ($Matches[1]) { $parts += "$($Matches[1]) Jahre" }
        if ($Matches[2]) { $parts += "$($Matches[2]) Monate" }
        if ($Matches[3]) { $parts += "$($Matches[3]) Tage" }
        if ($Matches[4]) { $parts += "$($Matches[4]) Stunden" }
        if ($Matches[5]) { $parts += "$($Matches[5]) Minuten" }
        if ($Matches[6]) { $parts += "$($Matches[6]) Sekunden" }
        if ($parts.Count -gt 0) { return ($parts -join " ") }
    }
    return $Duration
}

function Export-CsvUtf8Bom {
    param(
        [Parameter(Mandatory)]$Data,
        [Parameter(Mandatory)][string]$Path
    )
    $csv = ($Data | ConvertTo-Csv -NoTypeInformation -Delimiter ";")
    $utf8Bom = New-Object System.Text.UTF8Encoding $true
    [System.IO.File]::WriteAllLines($Path, $csv, $utf8Bom)
}

#endregion

#region ── Hauptprogramm ─────────────────────────────────────────────────────

Write-Log "=== Entra ID Gruppen Export gestartet ===" -Level "OK"

# ── Modul pruefen ────────────────────────────────────────────────────────────
if (-not (Get-Module -ListAvailable -Name Microsoft.Graph)) {
    Write-Log "Microsoft.Graph Modul nicht gefunden. Bitte installieren mit: Install-Module Microsoft.Graph" -Level "ERROR"
    exit 1
}

# ── Ausgabeordner erstellen ──────────────────────────────────────────────────
if (-not (Test-Path $OutputFolder)) {
    New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null
}
$OutputFolder = (Resolve-Path $OutputFolder).Path
Write-Log "Ausgabeordner: $OutputFolder"

# ── Verbindung herstellen (Least Privilege) ──────────────────────────────────
Write-Log "Verbinde mit Microsoft Graph (Least Privilege Scopes)..."
try {
    $connectParams = @{
        Scopes = @(
            "Group.Read.All",
            "PrivilegedEligibilitySchedule.Read.AzureADGroup",
            "PrivilegedAssignmentSchedule.Read.AzureADGroup",
            "RoleManagementPolicy.Read.AzureADGroup",
            "AccessReview.Read.All"
        )
        ErrorAction = "Stop"
    }
    if ($TenantId) { $connectParams.TenantId = $TenantId }
    Connect-MgGraph @connectParams

    $ctx = Get-MgContext
    Write-Log "Verbindung erfolgreich hergestellt." -Level "OK"
    Write-Log "Account  : $($ctx.Account)" -Level "OK"
    Write-Log "Tenant ID: $($ctx.TenantId)" -Level "OK"
}
catch {
    Write-Log "Fehler beim Verbinden mit Microsoft Graph: $_" -Level "ERROR"
    exit 1
}

# ── Alle Gruppen laden ──────────────────────────────────────────────────────
Write-Log "Lade alle Gruppen aus Entra ID..."
$groupUri  = "https://graph.microsoft.com/v1.0/groups?`$select=id,displayName,description,groupTypes,mailEnabled,securityEnabled,isAssignableToRole,membershipRuleProcessingState&`$top=999"
$allGroups = Get-GraphPagedResults -Uri $groupUri
Write-Log "$($allGroups.Count) Gruppen gefunden." -Level "OK"

# ── Access Review Definitionen vorladen ──────────────────────────────────────
Write-Log "Lade Access Review Definitionen..."
$accessReviewDefs = @()
try {
    $arUri = "https://graph.microsoft.com/v1.0/identityGovernance/accessReviews/definitions?`$select=id,displayName,scope,instanceEnumerationScope,status"
    $accessReviewDefs = Get-GraphPagedResults -Uri $arUri
    Write-Log "$($accessReviewDefs.Count) Access Review Definitionen geladen." -Level "OK"
}
catch {
    Write-Log "Access Reviews konnten nicht geladen werden (fehlende Berechtigung?): $_" -Level "WARN"
}

# ── Daten-Listen initialisieren ──────────────────────────────────────────────
$groupOverview    = [System.Collections.Generic.List[PSCustomObject]]::new()
$memberData       = [System.Collections.Generic.List[PSCustomObject]]::new()
$pimSettingsData  = [System.Collections.Generic.List[PSCustomObject]]::new()
$pimAssignData    = [System.Collections.Generic.List[PSCustomObject]]::new()

$totalGroups = $allGroups.Count
$counter     = 0

# ── Gruppen verarbeiten ─────────────────────────────────────────────────────
foreach ($group in $allGroups) {
    $counter++
    $pct = [math]::Round(($counter / $totalGroups) * 100)
    Write-Progress -Activity "Verarbeite Gruppen" -Status "$counter/$totalGroups - $($group.displayName)" -PercentComplete $pct
    Write-Log "[$counter/$totalGroups] $($group.displayName)"

    $groupId   = $group.id
    $groupName = $group.displayName
    $groupType = Get-GroupTypeName -Group $group
    $membershipType = Get-MembershipTypeName -Group $group
    $rolesAssignable = if ($group.isAssignableToRole) { "Ja" } else { "Nein" }

    # ── PIM-Status pruefen ───────────────────────────────────────────────────
    $pimEnabled = $false
    try {
        $eligUri  = "https://graph.microsoft.com/v1.0/identityGovernance/privilegedAccess/group/eligibilitySchedules?`$filter=groupId eq '${groupId}'&`$top=1"
        $eligResp = Invoke-MgGraphRequest -Uri $eligUri -Method GET -ErrorAction Stop
        if ($eligResp.value -and $eligResp.value.Count -gt 0) { $pimEnabled = $true }

        if (-not $pimEnabled) {
            $assignUri  = "https://graph.microsoft.com/v1.0/identityGovernance/privilegedAccess/group/assignmentSchedules?`$filter=groupId eq '${groupId}'&`$top=1"
            $assignResp = Invoke-MgGraphRequest -Uri $assignUri -Method GET -ErrorAction Stop
            if ($assignResp.value -and $assignResp.value.Count -gt 0) { $pimEnabled = $true }
        }

        if (-not $pimEnabled) {
            $policyCheckUri = "https://graph.microsoft.com/v1.0/policies/roleManagementPolicyAssignments?`$filter=scopeId eq '${groupId}' and scopeType eq 'Group'&`$top=1"
            $policyResp = Invoke-MgGraphRequest -Uri $policyCheckUri -Method GET -ErrorAction Stop
            if ($policyResp.value -and $policyResp.value.Count -gt 0) { $pimEnabled = $true }
        }
    }
    catch {
        # 403/404 = keine PIM-Konfiguration
    }

    # ── Gruppen-Uebersicht ───────────────────────────────────────────────────
    $groupOverview.Add([PSCustomObject]@{
        Gruppenname              = $groupName
        Beschreibung             = $group.description
        Gruppentyp               = $groupType
        Mitgliedschaftstyp       = $membershipType
        Rollenzuweisung_moeglich = $rolesAssignable
        PIM_aktiviert            = if ($pimEnabled) { "Ja" } else { "Nein" }
    })

    # ── Members und Owners laden ─────────────────────────────────────────────
    try {
        $membersUri = "https://graph.microsoft.com/v1.0/groups/${groupId}/members?`$select=id,displayName,userPrincipalName,mail&`$top=999"
        $members = Get-GraphPagedResults -Uri $membersUri
        foreach ($m in $members) {
            $objectType = ($m.'@odata.type' -replace '#microsoft\.graph\.', '')
            $upn = if ($m.userPrincipalName) { $m.userPrincipalName }
                   elseif ($m.mail)          { $m.mail }
                   else                      { $m.id }
            $memberData.Add([PSCustomObject]@{
                Gruppenname   = $groupName
                Anzeigename   = $m.displayName
                UPN           = $upn
                Objekttyp     = $objectType
                Rolle         = "Member"
            })
        }
    }
    catch {
        Write-Log "  Fehler beim Lesen der Members fuer ${groupName}: $_" -Level "WARN"
    }

    try {
        $ownersUri = "https://graph.microsoft.com/v1.0/groups/${groupId}/owners?`$select=id,displayName,userPrincipalName,mail&`$top=999"
        $owners = Get-GraphPagedResults -Uri $ownersUri
        foreach ($o in $owners) {
            $objectType = ($o.'@odata.type' -replace '#microsoft\.graph\.', '')
            $upn = if ($o.userPrincipalName) { $o.userPrincipalName }
                   elseif ($o.mail)          { $o.mail }
                   else                      { $o.id }
            $memberData.Add([PSCustomObject]@{
                Gruppenname   = $groupName
                Anzeigename   = $o.displayName
                UPN           = $upn
                Objekttyp     = $objectType
                Rolle         = "Besitzer"
            })
        }
    }
    catch {
        Write-Log "  Fehler beim Lesen der Owners fuer ${groupName}: $_" -Level "WARN"
    }

    # ── PIM-Details (nur fuer PIM-aktivierte Gruppen) ────────────────────────
    if ($pimEnabled) {
        Write-Log "  -> PIM aktiv. Lese Details..." -Level "OK"

        # ── Policy Assignments + Rules laden ─────────────────────────────────
        try {
            $paUri = "https://graph.microsoft.com/v1.0/policies/roleManagementPolicyAssignments?`$filter=scopeId eq '${groupId}' and scopeType eq 'Group'"
            $policyAssignments = Get-GraphPagedResults -Uri $paUri

            foreach ($pa in $policyAssignments) {
                $roleDefinition = $pa.roleDefinitionId  # "member" oder "owner"
                $policyId       = $pa.policyId

                # Policy mit Rules laden
                $policyUri = "https://graph.microsoft.com/v1.0/policies/roleManagementPolicies/${policyId}?`$expand=rules"
                $policy    = Invoke-MgGraphRequest -Uri $policyUri -Method GET -ErrorAction Stop
                $rules     = $policy.rules

                # Defaults
                $activationMaxDuration         = ""
                $activationMfa                 = "Nein"
                $activationJustification       = "Nein"
                $activationTicketing           = "Nein"
                $activationApprovalRequired    = "Nein"
                $activationApprovers           = ""
                $permanentEligibleAllowed      = "Nein"
                $eligibleExpirationAfter       = ""
                $permanentActiveAllowed        = "Nein"
                $activeExpirationAfter         = ""
                $activeAssignmentMfa           = "Nein"
                $activeAssignmentJustification = "Nein"

                foreach ($rule in $rules) {
                    switch ($rule.id) {
                        "Expiration_EndUser_Assignment" {
                            $activationMaxDuration = ConvertFrom-IsoDuration -Duration $rule.maximumDuration
                        }
                        "Enablement_EndUser_Assignment" {
                            $enabled = $rule.enabledRules
                            if ($enabled -contains "MultiFactorAuthentication") { $activationMfa = "Ja" }
                            if ($enabled -contains "Justification")             { $activationJustification = "Ja" }
                            if ($enabled -contains "Ticketing")                 { $activationTicketing = "Ja" }
                        }
                        "Approval_EndUser_Assignment" {
                            if ($rule.setting.isApprovalRequired) {
                                $activationApprovalRequired = "Ja"
                                $approverList = @()
                                foreach ($stage in $rule.setting.approvalStages) {
                                    foreach ($approver in $stage.primaryApprovers) {
                                        $approverName = if ($approver.description) { $approver.description }
                                                        elseif ($approver.userId)  { $approver.userId }
                                                        elseif ($approver.groupId) { $approver.groupId }
                                                        else                       { "Unbekannt" }
                                        $approverList += $approverName
                                    }
                                }
                                $activationApprovers = $approverList -join ", "
                            }
                        }
                        "Expiration_Admin_Eligibility" {
                            if (-not $rule.isExpirationRequired) {
                                $permanentEligibleAllowed = "Ja"
                            }
                            $eligibleExpirationAfter = ConvertFrom-IsoDuration -Duration $rule.maximumDuration
                        }
                        "Expiration_Admin_Assignment" {
                            if (-not $rule.isExpirationRequired) {
                                $permanentActiveAllowed = "Ja"
                            }
                            $activeExpirationAfter = ConvertFrom-IsoDuration -Duration $rule.maximumDuration
                        }
                        "Enablement_Admin_Assignment" {
                            $enabled = $rule.enabledRules
                            if ($enabled -contains "MultiFactorAuthentication") { $activeAssignmentMfa = "Ja" }
                            if ($enabled -contains "Justification")             { $activeAssignmentJustification = "Ja" }
                        }
                    }
                }

                # Access Review fuer diese Gruppe pruefen
                $arConfigured = "Nein"
                $arName       = ""
                foreach ($arDef in $accessReviewDefs) {
                    $scopeQuery = ""
                    if ($arDef.scope -and $arDef.scope.query) {
                        $scopeQuery = $arDef.scope.query
                    }
                    $enumQuery = ""
                    if ($arDef.instanceEnumerationScope -and $arDef.instanceEnumerationScope.query) {
                        $enumQuery = $arDef.instanceEnumerationScope.query
                    }
                    if ($scopeQuery -like "*${groupId}*" -or $enumQuery -like "*${groupId}*") {
                        $arConfigured = "Ja"
                        $arName = $arDef.displayName
                        break
                    }
                }

                $pimSettingsData.Add([PSCustomObject]@{
                    Gruppenname                              = $groupName
                    Rolle                                    = $roleDefinition
                    Aktivierung_Max_Dauer                    = $activationMaxDuration
                    Aktivierung_MFA_erforderlich              = $activationMfa
                    Aktivierung_Begruendung_erforderlich      = $activationJustification
                    Aktivierung_Ticketinfo_erforderlich       = $activationTicketing
                    Aktivierung_Genehmigung_erforderlich      = $activationApprovalRequired
                    Aktivierung_Genehmiger                   = $activationApprovers
                    Zuweisung_Permanent_Berechtigt_erlaubt   = $permanentEligibleAllowed
                    Zuweisung_Berechtigt_Ablauf_nach         = $eligibleExpirationAfter
                    Zuweisung_Permanent_Aktiv_erlaubt        = $permanentActiveAllowed
                    Zuweisung_Aktiv_Ablauf_nach              = $activeExpirationAfter
                    Zuweisung_Aktiv_MFA_erforderlich          = $activeAssignmentMfa
                    Zuweisung_Aktiv_Begruendung_erforderlich  = $activeAssignmentJustification
                    AccessReview_konfiguriert                = $arConfigured
                    AccessReview_Name                        = $arName
                })
            }
        }
        catch {
            Write-Log "  Fehler beim Lesen der PIM-Policies fuer ${groupName}: $_" -Level "WARN"
        }

        # ── Active Assignments ───────────────────────────────────────────────
        try {
            $activeUri   = "https://graph.microsoft.com/v1.0/identityGovernance/privilegedAccess/group/assignmentSchedules?`$filter=groupId eq '${groupId}'"
            $activeAssignments = Get-GraphPagedResults -Uri $activeUri
            foreach ($a in $activeAssignments) {
                $principalName = $a.principalId
                try {
                    $principalObj  = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/v1.0/directoryObjects/$($a.principalId)?`$select=displayName,userPrincipalName" -Method GET -ErrorAction Stop
                    $principalName = $principalObj.displayName
                    $principalUpn  = $principalObj.userPrincipalName
                }
                catch { $principalUpn = "" }

                $schedType = if ($a.scheduleInfo.expiration.type) { $a.scheduleInfo.expiration.type } else { "noExpiration" }
                $endDate   = if ($a.scheduleInfo.expiration.endDateTime) { $a.scheduleInfo.expiration.endDateTime } else { "-" }
                $startDate = if ($a.scheduleInfo.startDateTime) { $a.scheduleInfo.startDateTime } else { "-" }

                $pimAssignData.Add([PSCustomObject]@{
                    Gruppenname      = $groupName
                    Zuweisungstyp    = "Aktiv"
                    Anzeigename      = $principalName
                    UPN              = $principalUpn
                    Rolle            = $a.accessId
                    Status           = $a.status
                    Startdatum       = $startDate
                    Enddatum         = $endDate
                    Ablauftyp        = $schedType
                })
            }
        }
        catch {
            Write-Log "  Fehler beim Lesen der aktiven PIM-Zuweisungen fuer ${groupName}: $_" -Level "WARN"
        }

        # ── Eligible Assignments ─────────────────────────────────────────────
        try {
            $eligibleUri   = "https://graph.microsoft.com/v1.0/identityGovernance/privilegedAccess/group/eligibilitySchedules?`$filter=groupId eq '${groupId}'"
            $eligibleAssignments = Get-GraphPagedResults -Uri $eligibleUri
            foreach ($e in $eligibleAssignments) {
                $principalName = $e.principalId
                try {
                    $principalObj  = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/v1.0/directoryObjects/$($e.principalId)?`$select=displayName,userPrincipalName" -Method GET -ErrorAction Stop
                    $principalName = $principalObj.displayName
                    $principalUpn  = $principalObj.userPrincipalName
                }
                catch { $principalUpn = "" }

                $schedType = if ($e.scheduleInfo.expiration.type) { $e.scheduleInfo.expiration.type } else { "noExpiration" }
                $endDate   = if ($e.scheduleInfo.expiration.endDateTime) { $e.scheduleInfo.expiration.endDateTime } else { "-" }
                $startDate = if ($e.scheduleInfo.startDateTime) { $e.scheduleInfo.startDateTime } else { "-" }

                $pimAssignData.Add([PSCustomObject]@{
                    Gruppenname      = $groupName
                    Zuweisungstyp    = "Berechtigt"
                    Anzeigename      = $principalName
                    UPN              = $principalUpn
                    Rolle            = $e.accessId
                    Status           = $e.status
                    Startdatum       = $startDate
                    Enddatum         = $endDate
                    Ablauftyp        = $schedType
                })
            }
        }
        catch {
            Write-Log "  Fehler beim Lesen der berechtigten PIM-Zuweisungen fuer ${groupName}: $_" -Level "WARN"
        }
    }
}

Write-Progress -Activity "Verarbeite Gruppen" -Completed

# ── CSV-Export (UTF-8 mit BOM fuer korrekte Sonderzeichen) ───────────────────
Write-Log "Exportiere CSV-Dateien..."

$fileGroups  = Join-Path $OutputFolder "Gruppen_Uebersicht.csv"
$fileMembers = Join-Path $OutputFolder "Gruppen_Mitglieder.csv"
$filePimSet  = Join-Path $OutputFolder "PIM_Einstellungen.csv"
$filePimAss  = Join-Path $OutputFolder "PIM_Zuweisungen.csv"

try {
    if ($groupOverview.Count -gt 0) {
        Export-CsvUtf8Bom -Data $groupOverview -Path $fileGroups
        Write-Log "  $fileGroups ($($groupOverview.Count) Gruppen)" -Level "OK"
    }
    if ($memberData.Count -gt 0) {
        Export-CsvUtf8Bom -Data $memberData -Path $fileMembers
        Write-Log "  $fileMembers ($($memberData.Count) Eintraege)" -Level "OK"
    }
    if ($pimSettingsData.Count -gt 0) {
        Export-CsvUtf8Bom -Data $pimSettingsData -Path $filePimSet
        Write-Log "  $filePimSet ($($pimSettingsData.Count) Eintraege)" -Level "OK"
    }
    if ($pimAssignData.Count -gt 0) {
        Export-CsvUtf8Bom -Data $pimAssignData -Path $filePimAss
        Write-Log "  $filePimAss ($($pimAssignData.Count) Eintraege)" -Level "OK"
    }
    Write-Log "Export abgeschlossen!" -Level "OK"
}
catch {
    Write-Log "Fehler beim Schreiben der CSV-Dateien: $_" -Level "ERROR"
}

# ── Zusammenfassung ──────────────────────────────────────────────────────────
$pimGroupCount = ($groupOverview | Where-Object { $_.PIM_aktiviert -eq "Ja" }).Count
Write-Log "=== Zusammenfassung ===" -Level "OK"
Write-Log "  Gruppen gesamt     : $($groupOverview.Count)" -Level "OK"
Write-Log "  PIM-aktiviert      : $pimGroupCount" -Level "OK"
Write-Log "  Mitgliedschaften   : $($memberData.Count)" -Level "OK"
Write-Log "  PIM-Zuweisungen    : $($pimAssignData.Count)" -Level "OK"
Write-Log "  Ausgabeordner      : $OutputFolder" -Level "OK"

# ── Verbindung trennen ──────────────────────────────────────────────────────
Disconnect-MgGraph | Out-Null
Write-Log "Verbindung getrennt. Script beendet." -Level "OK"

#endregion
