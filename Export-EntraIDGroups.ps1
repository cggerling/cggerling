<#
.SYNOPSIS
    Exportiert Entra ID Daten (Gruppen, Benutzer, Rollen) in CSV-Dateien.
    Interaktive Modulauswahl beim Start - nur gewaehlte Bereiche werden abgefragt.

.DESCRIPTION
    Das Script verbindet sich mit Microsoft Graph (Least Privilege / Zero Trust).
    Vor dem Start wird eine Modulauswahl angezeigt:

    [1] Gruppen-Uebersicht         -> Gruppen_Uebersicht.csv
    [2] Gruppen-Mitgliedschaften   -> Gruppen_Mitglieder.csv
    [3] PIM fuer Gruppen           -> PIM_Einstellungen.csv, PIM_Zuweisungen.csv
    [4] Benutzer                   -> Benutzer.csv
    [5] Directory-Rollenzuweisungen-> Rollen_Aktiv.csv, Rollen_Eligible.csv
    [0] Alle Module exportieren

    Berechtigungen werden dynamisch anhand der Auswahl angefordert (Least Privilege).

.PARAMETER OutputFolder
    Ordner fuer die Export-Dateien. Standard: .\EntraID_Export_<Datum>

.PARAMETER TenantId
    Optionale Tenant-ID fuer Multi-Tenant-Umgebungen.

.EXAMPLE
    .\Export-EntraIDGroups.ps1
    .\Export-EntraIDGroups.ps1 -TenantId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
    .\Export-EntraIDGroups.ps1 -OutputFolder "C:\Reports\EntraExport"

.NOTES
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

    if ($gt -contains "Unified")                { return "Microsoft 365" }
    elseif ($secEnabled -and $mailEnabled)      { return "Mail-enabled Security" }
    elseif ($secEnabled -and -not $mailEnabled) { return "Security" }
    elseif ($mailEnabled -and -not $secEnabled) { return "Distribution" }
    else                                        { return "Unbekannt" }
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

#region ── Moduldefinitionen ─────────────────────────────────────────────────
# Neue Module koennen hier einfach als weiterer Eintrag ergaenzt werden.
# Jedes Modul definiert seinen Namen, eine Beschreibung, einen internen Key
# und die minimal benoetigten Graph-Berechtigungen (Least Privilege).

$script:moduleDefinitions = @(
    [PSCustomObject]@{
        Id     = 1
        Name   = "Gruppen-Uebersicht"
        Desc   = "Name, Typ, Mitgliedschaftstyp, PIM-Status"
        Key    = "Gruppen"
        Scopes = @("Group.Read.All")
    }
    [PSCustomObject]@{
        Id     = 2
        Name   = "Gruppen-Mitgliedschaften"
        Desc   = "Members und Owners aller Gruppen"
        Key    = "Mitglieder"
        Scopes = @("Group.Read.All")
    }
    [PSCustomObject]@{
        Id     = 3
        Name   = "PIM fuer Gruppen"
        Desc   = "Settings, Zuweisungen, Eligible Members, Access Reviews"
        Key    = "PIM"
        Scopes = @(
            "Group.Read.All",
            "PrivilegedEligibilitySchedule.Read.AzureADGroup",
            "PrivilegedAssignmentSchedule.Read.AzureADGroup",
            "RoleManagementPolicy.Read.AzureADGroup",
            "AccessReview.Read.All"
        )
    }
    [PSCustomObject]@{
        Id     = 4
        Name   = "Benutzer"
        Desc   = "DisplayName, Vorname, Nachname, UPN, Typ, Status, OnPrem-Sync"
        Key    = "Benutzer"
        Scopes = @("User.Read.All")
    }
    [PSCustomObject]@{
        Id     = 5
        Name   = "Directory-Rollenzuweisungen"
        Desc   = "Aktive (statisch + PIM) und berechtigte Admin-Rollen"
        Key    = "Rollen"
        Scopes = @("User.Read.All", "RoleManagement.Read.Directory")
    }
    # ── Hier weitere Module ergaenzen ──────────────────────────────────────
    # [PSCustomObject]@{
    #     Id     = 6
    #     Name   = "Conditional Access Policies"
    #     Desc   = "Alle CA-Richtlinien mit Zuweisungen"
    #     Key    = "ConditionalAccess"
    #     Scopes = @("Policy.Read.All")
    # }
)

function Show-ModuleMenu {
    param([PSCustomObject[]]$Modules)

    $border = "=" * 66
    Write-Host ""
    Write-Host "  $border" -ForegroundColor Cyan
    Write-Host "   Entra ID Export - Modulauswahl" -ForegroundColor Cyan
    Write-Host "  $border" -ForegroundColor Cyan

    foreach ($m in $Modules) {
        Write-Host ("   [{0}]  {1,-35} {2}" -f $m.Id, $m.Name, $m.Desc) -ForegroundColor White
    }

    Write-Host "  $border" -ForegroundColor Cyan
    Write-Host "   [0]  Alle Module exportieren" -ForegroundColor Green
    Write-Host "  $border" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "   Mehrfachauswahl moeglich, Nummern kommagetrennt eingeben." -ForegroundColor Gray
    Write-Host "   Beispiele: '1'  |  '1,3,5'  |  '0' fuer alle" -ForegroundColor Gray
    Write-Host ""

    do {
        $rawInput = Read-Host "  Auswahl"
        $rawInput = $rawInput.Trim()
        $isValid  = $rawInput -match '^[0-9][0-9,\s]*$'
        if (-not $isValid) {
            Write-Host "  Ungueltige Eingabe. Bitte Nummern kommagetrennt eingeben (z.B. 1,3)." -ForegroundColor Yellow
        }
    } while (-not $isValid)

    $selected = $rawInput -split '[,\s]+' |
                Where-Object { $_ -ne '' } |
                ForEach-Object { [int]$_ } |
                Sort-Object -Unique

    if ($selected -contains 0) {
        Write-Host "  -> Alle Module gewaehlt." -ForegroundColor Green
        return ($Modules | Select-Object -ExpandProperty Key)
    }

    $validIds      = $Modules | Select-Object -ExpandProperty Id
    $unknownIds    = $selected | Where-Object { $validIds -notcontains $_ }
    if ($unknownIds) {
        Write-Host "  Unbekannte Modul-IDs werden ignoriert: $($unknownIds -join ', ')" -ForegroundColor Yellow
    }

    $chosenModules = $Modules | Where-Object { $selected -contains $_.Id }
    if (-not $chosenModules) {
        Write-Host "  Keine gueltigen Module gewaehlt. Bitte erneut eingeben." -ForegroundColor Yellow
        return Show-ModuleMenu -Modules $Modules
    }

    Write-Host "  -> Gewaehlt: $($chosenModules.Name -join ', ')" -ForegroundColor Green
    return ($chosenModules | Select-Object -ExpandProperty Key)
}

#endregion

#region ── Hauptprogramm ─────────────────────────────────────────────────────

Write-Log "=== Entra ID Export gestartet ===" -Level "OK"

# ── Modul pruefen ────────────────────────────────────────────────────────────
if (-not (Get-Module -ListAvailable -Name Microsoft.Graph)) {
    Write-Log "Microsoft.Graph Modul nicht gefunden. Bitte installieren: Install-Module Microsoft.Graph" -Level "ERROR"
    exit 1
}

# ── Ausgabeordner erstellen ──────────────────────────────────────────────────
if (-not (Test-Path $OutputFolder)) {
    New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null
}
$OutputFolder = (Resolve-Path $OutputFolder).Path

# ── Modulauswahl ─────────────────────────────────────────────────────────────
$selectedKeys  = Show-ModuleMenu -Modules $script:moduleDefinitions

$runGruppen    = $selectedKeys -contains "Gruppen"
$runMitglieder = $selectedKeys -contains "Mitglieder"
$runPIM        = $selectedKeys -contains "PIM"
$runBenutzer   = $selectedKeys -contains "Benutzer"
$runRollen     = $selectedKeys -contains "Rollen"

# Gruppen muessen fuer Mitglieder und PIM mitgeladen werden
$loadGroups    = $runGruppen -or $runMitglieder -or $runPIM

# ── Berechtigungen dynamisch aus Modulauswahl ableiten ───────────────────────
$requiredScopes = $script:moduleDefinitions |
    Where-Object { $selectedKeys -contains $_.Key } |
    Select-Object -ExpandProperty Scopes |
    ForEach-Object { $_ } |
    Sort-Object -Unique

Write-Log "Ausgabeordner : $OutputFolder"
Write-Log "Benoetigte Berechtigungen:"
$requiredScopes | ForEach-Object { Write-Log "  - $_" }

# ── Verbindung herstellen ────────────────────────────────────────────────────
Write-Log "Verbinde mit Microsoft Graph..."
try {
    $connectParams = @{
        Scopes      = $requiredScopes
        ErrorAction = "Stop"
    }
    if ($TenantId) { $connectParams.TenantId = $TenantId }
    Connect-MgGraph @connectParams

    $ctx = Get-MgContext
    Write-Log "Verbunden." -Level "OK"
    Write-Log "Account  : $($ctx.Account)" -Level "OK"
    Write-Log "Tenant ID: $($ctx.TenantId)" -Level "OK"
}
catch {
    Write-Log "Fehler beim Verbinden mit Microsoft Graph: $_" -Level "ERROR"
    exit 1
}

# ── Daten-Listen initialisieren ──────────────────────────────────────────────
$groupOverview   = [System.Collections.Generic.List[PSCustomObject]]::new()
$memberData      = [System.Collections.Generic.List[PSCustomObject]]::new()
$pimSettingsData = [System.Collections.Generic.List[PSCustomObject]]::new()
$pimAssignData   = [System.Collections.Generic.List[PSCustomObject]]::new()
$userData        = [System.Collections.Generic.List[PSCustomObject]]::new()
$rollenAktivData = [System.Collections.Generic.List[PSCustomObject]]::new()
$rollenEligData  = [System.Collections.Generic.List[PSCustomObject]]::new()

#endregion

#region ── Modul: Gruppen / Mitglieder / PIM ─────────────────────────────────

if ($loadGroups) {
    # ── Alle Gruppen laden ───────────────────────────────────────────────────
    Write-Log "Lade alle Gruppen aus Entra ID..."
    $groupUri  = "https://graph.microsoft.com/v1.0/groups?`$select=id,displayName,description,groupTypes,mailEnabled,securityEnabled,isAssignableToRole,membershipRuleProcessingState&`$top=999"
    $allGroups = Get-GraphPagedResults -Uri $groupUri
    Write-Log "$($allGroups.Count) Gruppen gefunden." -Level "OK"

    # ── Access Reviews vorladen (nur fuer PIM benoetigt) ─────────────────────
    $accessReviewDefs = @()
    if ($runPIM) {
        Write-Log "Lade Access Review Definitionen..."
        try {
            $arUri = "https://graph.microsoft.com/v1.0/identityGovernance/accessReviews/definitions?`$select=id,displayName,scope,instanceEnumerationScope,status"
            $accessReviewDefs = Get-GraphPagedResults -Uri $arUri
            Write-Log "$($accessReviewDefs.Count) Access Review Definitionen geladen." -Level "OK"
        }
        catch {
            Write-Log "Access Reviews nicht verfuegbar (Berechtigung?): $_" -Level "WARN"
        }
    }

    # ── Gruppen verarbeiten ──────────────────────────────────────────────────
    $totalGroups = $allGroups.Count
    $counter     = 0

    foreach ($group in $allGroups) {
        $counter++
        $pct = [math]::Round(($counter / $totalGroups) * 100)
        Write-Progress -Activity "Verarbeite Gruppen" -Status "$counter/$totalGroups - $($group.displayName)" -PercentComplete $pct

        $groupId        = $group.id
        $groupName      = $group.displayName
        $groupType      = Get-GroupTypeName -Group $group
        $membershipType = Get-MembershipTypeName -Group $group

        # ── PIM-Status pruefen (benoetigt fuer Gruppen-Uebersicht + PIM) ────
        $pimEnabled = $false
        if ($runGruppen -or $runPIM) {
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
                    $policyResp     = Invoke-MgGraphRequest -Uri $policyCheckUri -Method GET -ErrorAction Stop
                    if ($policyResp.value -and $policyResp.value.Count -gt 0) { $pimEnabled = $true }
                }
            }
            catch { <# 403/404 = kein PIM #> }
        }

        # ── Modul: Gruppen-Uebersicht ────────────────────────────────────────
        if ($runGruppen) {
            $groupOverview.Add([PSCustomObject]@{
                Gruppenname              = $groupName
                Beschreibung             = $group.description
                Gruppentyp               = $groupType
                Mitgliedschaftstyp       = $membershipType
                Rollenzuweisung_moeglich = if ($group.isAssignableToRole) { "Ja" } else { "Nein" }
                PIM_aktiviert            = if ($pimEnabled) { "Ja" } else { "Nein" }
            })
        }

        # ── Modul: Gruppen-Mitgliedschaften ──────────────────────────────────
        if ($runMitglieder) {
            try {
                $membersUri = "https://graph.microsoft.com/v1.0/groups/${groupId}/members?`$select=id,displayName,userPrincipalName,mail&`$top=999"
                $members    = Get-GraphPagedResults -Uri $membersUri
                foreach ($m in $members) {
                    $objectType = ($m.'@odata.type' -replace '#microsoft\.graph\.', '')
                    $upn = if ($m.userPrincipalName) { $m.userPrincipalName }
                           elseif ($m.mail)          { $m.mail }
                           else                      { $m.id }
                    $memberData.Add([PSCustomObject]@{
                        Gruppenname = $groupName
                        Anzeigename = $m.displayName
                        UPN         = $upn
                        Objekttyp   = $objectType
                        Rolle       = "Member"
                    })
                }
            }
            catch { Write-Log "  Fehler Members fuer ${groupName}: $_" -Level "WARN" }

            try {
                $ownersUri = "https://graph.microsoft.com/v1.0/groups/${groupId}/owners?`$select=id,displayName,userPrincipalName,mail&`$top=999"
                $owners    = Get-GraphPagedResults -Uri $ownersUri
                foreach ($o in $owners) {
                    $objectType = ($o.'@odata.type' -replace '#microsoft\.graph\.', '')
                    $upn = if ($o.userPrincipalName) { $o.userPrincipalName }
                           elseif ($o.mail)          { $o.mail }
                           else                      { $o.id }
                    $memberData.Add([PSCustomObject]@{
                        Gruppenname = $groupName
                        Anzeigename = $o.displayName
                        UPN         = $upn
                        Objekttyp   = $objectType
                        Rolle       = "Besitzer"
                    })
                }
            }
            catch { Write-Log "  Fehler Owners fuer ${groupName}: $_" -Level "WARN" }
        }

        # ── Modul: PIM fuer Gruppen ──────────────────────────────────────────
        if ($runPIM -and $pimEnabled) {
            Write-Log "  -> PIM aktiv. Lese Details..." -Level "OK"

            # Policy Assignments + Rules
            try {
                $paUri             = "https://graph.microsoft.com/v1.0/policies/roleManagementPolicyAssignments?`$filter=scopeId eq '${groupId}' and scopeType eq 'Group'"
                $policyAssignments = Get-GraphPagedResults -Uri $paUri

                foreach ($pa in $policyAssignments) {
                    $roleDefinition = $pa.roleDefinitionId
                    $policyId       = $pa.policyId

                    $policyUri = "https://graph.microsoft.com/v1.0/policies/roleManagementPolicies/${policyId}?`$expand=rules"
                    $policy    = Invoke-MgGraphRequest -Uri $policyUri -Method GET -ErrorAction Stop
                    $rules     = $policy.rules

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
                                            $approverList += if ($approver.description) { $approver.description }
                                                             elseif ($approver.userId)   { $approver.userId }
                                                             elseif ($approver.groupId)  { $approver.groupId }
                                                             else                        { "Unbekannt" }
                                        }
                                    }
                                    $activationApprovers = $approverList -join ", "
                                }
                            }
                            "Expiration_Admin_Eligibility" {
                                $permanentEligibleAllowed = if (-not $rule.isExpirationRequired) { "Ja" } else { "Nein" }
                                $eligibleExpirationAfter  = ConvertFrom-IsoDuration -Duration $rule.maximumDuration
                            }
                            "Expiration_Admin_Assignment" {
                                $permanentActiveAllowed = if (-not $rule.isExpirationRequired) { "Ja" } else { "Nein" }
                                $activeExpirationAfter  = ConvertFrom-IsoDuration -Duration $rule.maximumDuration
                            }
                            "Enablement_Admin_Assignment" {
                                $enabled = $rule.enabledRules
                                if ($enabled -contains "MultiFactorAuthentication") { $activeAssignmentMfa = "Ja" }
                                if ($enabled -contains "Justification")             { $activeAssignmentJustification = "Ja" }
                            }
                        }
                    }

                    # Access Review pruefen
                    $arConfigured = "Nein"
                    $arName       = ""
                    foreach ($arDef in $accessReviewDefs) {
                        $sq = if ($arDef.scope -and $arDef.scope.query) { $arDef.scope.query } else { "" }
                        $eq = if ($arDef.instanceEnumerationScope -and $arDef.instanceEnumerationScope.query) { $arDef.instanceEnumerationScope.query } else { "" }
                        if ($sq -like "*${groupId}*" -or $eq -like "*${groupId}*") {
                            $arConfigured = "Ja"
                            $arName       = $arDef.displayName
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
            catch { Write-Log "  Fehler PIM-Policies fuer ${groupName}: $_" -Level "WARN" }

            # Active Assignments
            try {
                $activeAssignments = Get-GraphPagedResults -Uri "https://graph.microsoft.com/v1.0/identityGovernance/privilegedAccess/group/assignmentSchedules?`$filter=groupId eq '${groupId}'"
                foreach ($a in $activeAssignments) {
                    $principalName = $a.principalId
                    $principalUpn  = ""
                    try {
                        $po = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/v1.0/directoryObjects/$($a.principalId)?`$select=displayName,userPrincipalName" -Method GET -ErrorAction Stop
                        $principalName = $po.displayName
                        $principalUpn  = $po.userPrincipalName
                    } catch {}
                    $pimAssignData.Add([PSCustomObject]@{
                        Gruppenname   = $groupName
                        Zuweisungstyp = "Aktiv"
                        Anzeigename   = $principalName
                        UPN           = $principalUpn
                        Rolle         = $a.accessId
                        Status        = $a.status
                        Startdatum    = if ($a.scheduleInfo.startDateTime) { $a.scheduleInfo.startDateTime } else { "-" }
                        Enddatum      = if ($a.scheduleInfo.expiration.endDateTime) { $a.scheduleInfo.expiration.endDateTime } else { "-" }
                        Ablauftyp     = if ($a.scheduleInfo.expiration.type) { $a.scheduleInfo.expiration.type } else { "noExpiration" }
                    })
                }
            }
            catch { Write-Log "  Fehler aktive PIM-Zuweisungen fuer ${groupName}: $_" -Level "WARN" }

            # Eligible Assignments
            try {
                $eligibleAssignments = Get-GraphPagedResults -Uri "https://graph.microsoft.com/v1.0/identityGovernance/privilegedAccess/group/eligibilitySchedules?`$filter=groupId eq '${groupId}'"
                foreach ($e in $eligibleAssignments) {
                    $principalName = $e.principalId
                    $principalUpn  = ""
                    try {
                        $po = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/v1.0/directoryObjects/$($e.principalId)?`$select=displayName,userPrincipalName" -Method GET -ErrorAction Stop
                        $principalName = $po.displayName
                        $principalUpn  = $po.userPrincipalName
                    } catch {}
                    $pimAssignData.Add([PSCustomObject]@{
                        Gruppenname   = $groupName
                        Zuweisungstyp = "Berechtigt"
                        Anzeigename   = $principalName
                        UPN           = $principalUpn
                        Rolle         = $e.accessId
                        Status        = $e.status
                        Startdatum    = if ($e.scheduleInfo.startDateTime) { $e.scheduleInfo.startDateTime } else { "-" }
                        Enddatum      = if ($e.scheduleInfo.expiration.endDateTime) { $e.scheduleInfo.expiration.endDateTime } else { "-" }
                        Ablauftyp     = if ($e.scheduleInfo.expiration.type) { $e.scheduleInfo.expiration.type } else { "noExpiration" }
                    })
                }
            }
            catch { Write-Log "  Fehler berechtigte PIM-Zuweisungen fuer ${groupName}: $_" -Level "WARN" }
        }
    }

    Write-Progress -Activity "Verarbeite Gruppen" -Completed
}

#endregion

#region ── Modul: Benutzer ───────────────────────────────────────────────────

if ($runBenutzer) {
    Write-Log "Lade alle Benutzer aus Entra ID..."
    try {
        $userUri  = "https://graph.microsoft.com/v1.0/users?`$select=displayName,givenName,surname,userPrincipalName,userType,accountEnabled,onPremisesSyncEnabled&`$top=999"
        $allUsers = Get-GraphPagedResults -Uri $userUri
        Write-Log "$($allUsers.Count) Benutzer gefunden." -Level "OK"

        foreach ($u in $allUsers) {
            $userData.Add([PSCustomObject]@{
                Anzeigename       = $u.displayName
                Vorname           = $u.givenName
                Nachname          = $u.surname
                UserPrincipalName = $u.userPrincipalName
                Benutzertyp       = if ($u.userType) { $u.userType } else { "Member" }
                Konto_aktiv       = if ($u.accountEnabled) { "Ja" } else { "Nein" }
                OnPremises_Sync   = if ($u.onPremisesSyncEnabled) { "Ja" } else { "Nein" }
            })
        }
    }
    catch { Write-Log "Fehler beim Laden der Benutzer: $_" -Level "ERROR" }
}

#endregion

#region ── Modul: Directory-Rollenzuweisungen ────────────────────────────────

if ($runRollen) {
    # Rollendefinitionen laden (ID -> Name)
    Write-Log "Lade Directory-Rollendefinitionen..."
    $roleDefMap = @{}
    try {
        $roleDefs = Get-GraphPagedResults -Uri "https://graph.microsoft.com/v1.0/roleManagement/directory/roleDefinitions?`$select=id,displayName"
        foreach ($rd in $roleDefs) { $roleDefMap[$rd.id] = $rd.displayName }
        Write-Log "$($roleDefs.Count) Rollendefinitionen geladen." -Level "OK"
    }
    catch { Write-Log "Fehler beim Laden der Rollendefinitionen: $_" -Level "WARN" }

    # Aktive Rollenzuweisungen
    Write-Log "Lade aktive Directory-Rollenzuweisungen..."
    try {
        $allAssign   = Get-GraphPagedResults -Uri "https://graph.microsoft.com/v1.0/roleManagement/directory/roleAssignments?`$expand=principal(`$select=displayName,userPrincipalName,id)&`$top=999"
        $schedules   = Get-GraphPagedResults -Uri "https://graph.microsoft.com/v1.0/roleManagement/directory/roleAssignmentSchedules?`$select=principalId,roleDefinitionId&`$top=999"
        $schedLookup = @{}
        foreach ($s in $schedules) { $schedLookup["$($s.principalId)|$($s.roleDefinitionId)"] = $true }

        Write-Log "$($allAssign.Count) aktive Rollenzuweisungen gefunden." -Level "OK"

        foreach ($a in $allAssign) {
            $principalType = $a.principal.'@odata.type' -replace '#microsoft\.graph\.', ''
            if ($principalType -ne 'user') { continue }

            $roleName   = if ($roleDefMap[$a.roleDefinitionId]) { $roleDefMap[$a.roleDefinitionId] } else { $a.roleDefinitionId }
            $isPim      = $schedLookup["$($a.principalId)|$($a.roleDefinitionId)"]
            $assignType = if ($isPim) { "PIM (aktiv)" } else { "Statisch (permanent)" }
            $scope      = if ($a.directoryScopeId -eq "/") { "Tenant (global)" } else { $a.directoryScopeId }

            $rollenAktivData.Add([PSCustomObject]@{
                Anzeigename   = $a.principal.displayName
                UPN           = $a.principal.userPrincipalName
                Rollenname    = $roleName
                Zuweisungstyp = $assignType
                Scope         = $scope
            })
        }
    }
    catch { Write-Log "Fehler beim Laden der aktiven Rollenzuweisungen: $_" -Level "ERROR" }

    # Berechtigte (Eligible) Rollenzuweisungen
    Write-Log "Lade berechtigte (eligible) Directory-Rollenzuweisungen..."
    try {
        $allElig = Get-GraphPagedResults -Uri "https://graph.microsoft.com/v1.0/roleManagement/directory/roleEligibilitySchedules?`$expand=principal&`$top=999"
        Write-Log "$($allElig.Count) berechtigte Rollenzuweisungen gefunden." -Level "OK"

        foreach ($e in $allElig) {
            $principalType = $e.principal.'@odata.type' -replace '#microsoft\.graph\.', ''
            if ($principalType -ne 'user') { continue }

            $roleName  = if ($roleDefMap[$e.roleDefinitionId]) { $roleDefMap[$e.roleDefinitionId] } else { $e.roleDefinitionId }
            $scope     = if ($e.directoryScopeId -eq "/") { "Tenant (global)" } else { $e.directoryScopeId }
            $schedType = if ($e.scheduleInfo.expiration.type -and $e.scheduleInfo.expiration.type -ne 'noExpiration') { "Zeitgebunden" } else { "Permanent" }

            $rollenEligData.Add([PSCustomObject]@{
                Anzeigename   = $e.principal.displayName
                UPN           = $e.principal.userPrincipalName
                Rollenname    = $roleName
                Ablauftyp     = $schedType
                Scope         = $scope
                Startdatum    = if ($e.scheduleInfo.startDateTime) { $e.scheduleInfo.startDateTime } else { "-" }
                Enddatum      = if ($e.scheduleInfo.expiration.endDateTime) { $e.scheduleInfo.expiration.endDateTime } else { "-" }
                Status        = $e.status
            })
        }
    }
    catch { Write-Log "Fehler beim Laden der berechtigten Rollenzuweisungen: $_" -Level "ERROR" }
}

#endregion

#region ── CSV-Export ────────────────────────────────────────────────────────

Write-Log "Exportiere CSV-Dateien (UTF-8 mit BOM)..."

$exports = @(
    @{ Run = $runGruppen;    Data = $groupOverview;   File = "Gruppen_Uebersicht.csv";   Label = "Gruppen" }
    @{ Run = $runMitglieder; Data = $memberData;      File = "Gruppen_Mitglieder.csv";   Label = "Gruppen-Mitgliedschaften" }
    @{ Run = $runPIM;        Data = $pimSettingsData; File = "PIM_Einstellungen.csv";    Label = "PIM-Einstellungen" }
    @{ Run = $runPIM;        Data = $pimAssignData;   File = "PIM_Zuweisungen.csv";      Label = "PIM-Zuweisungen" }
    @{ Run = $runBenutzer;   Data = $userData;        File = "Benutzer.csv";             Label = "Benutzer" }
    @{ Run = $runRollen;     Data = $rollenAktivData; File = "Rollen_Aktiv.csv";         Label = "Rollen aktiv" }
    @{ Run = $runRollen;     Data = $rollenEligData;  File = "Rollen_Eligible.csv";      Label = "Rollen eligible" }
)

foreach ($export in $exports) {
    if (-not $export.Run) { continue }
    if ($export.Data.Count -eq 0) {
        Write-Log "  $($export.Label): keine Daten gefunden." -Level "WARN"
        continue
    }
    try {
        $filePath = Join-Path $OutputFolder $export.File
        Export-CsvUtf8Bom -Data $export.Data -Path $filePath
        Write-Log "  $($export.File) ($($export.Data.Count) Eintraege)" -Level "OK"
    }
    catch {
        Write-Log "  Fehler beim Schreiben von $($export.File): $_" -Level "ERROR"
    }
}

Write-Log "Export abgeschlossen!" -Level "OK"

#endregion

#region ── Zusammenfassung ───────────────────────────────────────────────────

Write-Log "=== Zusammenfassung ===" -Level "OK"
if ($runGruppen)    { Write-Log "  Gruppen               : $($groupOverview.Count)" -Level "OK" }
if ($runMitglieder) { Write-Log "  Mitgliedschaften      : $($memberData.Count)" -Level "OK" }
if ($runPIM) {
    $pimCount = ($groupOverview | Where-Object { $_.PIM_aktiviert -eq "Ja" }).Count
    Write-Log "  PIM-Gruppen           : $pimCount" -Level "OK"
    Write-Log "  PIM-Zuweisungen       : $($pimAssignData.Count)" -Level "OK"
}
if ($runBenutzer)   { Write-Log "  Benutzer              : $($userData.Count)" -Level "OK" }
if ($runRollen) {
    $staticCount = ($rollenAktivData | Where-Object { $_.Zuweisungstyp -like "Statisch*" }).Count
    $pimRolCount = ($rollenAktivData | Where-Object { $_.Zuweisungstyp -like "PIM*" }).Count
    Write-Log "  Rollen aktiv (statisch): $staticCount" -Level "OK"
    Write-Log "  Rollen aktiv (PIM)     : $pimRolCount" -Level "OK"
    Write-Log "  Rollen eligible (PIM)  : $($rollenEligData.Count)" -Level "OK"
}
Write-Log "  Ausgabeordner         : $OutputFolder" -Level "OK"

Disconnect-MgGraph | Out-Null
Write-Log "Verbindung getrennt. Script beendet." -Level "OK"

#endregion
