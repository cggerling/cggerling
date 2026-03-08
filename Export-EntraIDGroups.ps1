<#
.SYNOPSIS
    Exportiert alle Entra ID Gruppen inkl. Members und PIM-for-Groups-Informationen in eine CSV-Datei.

.DESCRIPTION
    Das Script verbindet sich mit Microsoft Graph und liest alle Gruppen aus der Entra ID aus.
    Pro Gruppe werden folgende Werte ermittelt:
      - Gruppenname
      - Typ (Security, Microsoft 365, Distribution, Mail-enabled Security)
      - Members (DisplayName | UPN)
      - PIM for Groups aktiviert (Ja/Nein)
      - PIM-Einstellungen (Policy)
      - PIM Role Assignments (aktive Zuweisungen)
      - PIM Eligible Members (berechtigte Mitglieder)

.PARAMETER OutputPath
    Pfad zur Ausgabe-CSV-Datei. Standard: .\EntraID_Groups_Export_<Datum>.csv

.PARAMETER BatchSize
    Anzahl der Gruppen, die parallel verarbeitet werden (Standard: 20).

.EXAMPLE
    .\Export-EntraIDGroups.ps1
    .\Export-EntraIDGroups.ps1 -OutputPath "C:\Reports\groups.csv"

.NOTES
    Benötigte Microsoft Graph Berechtigungen (Application oder Delegated):
      - Group.Read.All
      - GroupMember.Read.All
      - PrivilegedAccess.Read.AzureADGroup
      - RoleManagementPolicy.Read.AzureADGroup

    Voraussetzung: Microsoft.Graph PowerShell SDK (Install-Module Microsoft.Graph)
#>

[CmdletBinding()]
param(
    [string]$OutputPath = ".\EntraID_Groups_Export_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv",
    [int]$BatchSize = 20
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
    <#
    .SYNOPSIS Liest alle Seiten einer paginierten Graph-Antwort aus.
    #>
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

function Get-GroupType {
    <#
    .SYNOPSIS Ermittelt den lesbaren Gruppentyp anhand der Graph-Eigenschaften.
    #>
    param($Group)
    $gt = $Group.groupTypes
    $mailEnabled  = $Group.mailEnabled
    $secEnabled   = $Group.securityEnabled

    if ($gt -contains "Unified") {
        return "Microsoft 365"
    }
    elseif ($secEnabled -and $mailEnabled) {
        return "Mail-enabled Security"
    }
    elseif ($secEnabled -and -not $mailEnabled) {
        return "Security"
    }
    elseif ($mailEnabled -and -not $secEnabled) {
        return "Distribution"
    }
    else {
        return "Unbekannt"
    }
}

function Get-GroupMembers {
    <#
    .SYNOPSIS Liest alle direkten Member einer Gruppe und gibt sie als String zurück.
    #>
    param([string]$GroupId)
    try {
        $uri     = "https://graph.microsoft.com/v1.0/groups/$GroupId/members?`$select=displayName,userPrincipalName,mail,id&`$top=999"
        $members = Get-GraphPagedResults -Uri $uri
        if (-not $members) { return "" }

        $memberStrings = foreach ($m in $members) {
            $upn = if ($m.userPrincipalName) { $m.userPrincipalName }
                   elseif ($m.mail)          { $m.mail }
                   else                      { $m.id }
            "$($m.displayName) ($upn)"
        }
        return ($memberStrings -join " | ")
    }
    catch {
        Write-Log "Fehler beim Lesen der Member für Gruppe $GroupId: $_" -Level "WARN"
        return "FEHLER"
    }
}

function Get-PimStatus {
    <#
    .SYNOPSIS
        Prüft ob PIM for Groups für eine Gruppe aktiv ist.
        Eine Gruppe gilt als PIM-fähig, wenn mindestens ein aktives oder berechtigtes
        Assignment-Schedule vorhanden ist ODER eine RoleManagementPolicy existiert.
    #>
    param([string]$GroupId)
    try {
        # Prüfe auf Eligible Schedules
        $eligUri = "https://graph.microsoft.com/v1.0/identityGovernance/privilegedAccess/group/eligibilitySchedules?`$filter=groupId eq '$GroupId'&`$top=1"
        $eligResp = Invoke-MgGraphRequest -Uri $eligUri -Method GET -ErrorAction Stop
        if ($eligResp.value -and $eligResp.value.Count -gt 0) { return $true }

        # Prüfe auf Active Assignment Schedules
        $assignUri = "https://graph.microsoft.com/v1.0/identityGovernance/privilegedAccess/group/assignmentSchedules?`$filter=groupId eq '$GroupId'&`$top=1"
        $assignResp = Invoke-MgGraphRequest -Uri $assignUri -Method GET -ErrorAction Stop
        if ($assignResp.value -and $assignResp.value.Count -gt 0) { return $true }

        return $false
    }
    catch {
        # 403 oder 404 = keine PIM-Konfiguration für diese Gruppe
        return $false
    }
}

function Get-PimSettings {
    <#
    .SYNOPSIS Liest die RoleManagementPolicy-Einstellungen für eine PIM-Gruppe.
    #>
    param([string]$GroupId)
    try {
        $uri      = "https://graph.microsoft.com/v1.0/policies/roleManagementPolicies?`$filter=scopeId eq '$GroupId' and scopeType eq 'Group'"
        $policies = Get-GraphPagedResults -Uri $uri
        if (-not $policies) { return "" }

        $settingLines = foreach ($policy in $policies) {
            $rulesSummary = @()
            if ($policy.rules) {
                foreach ($rule in $policy.rules) {
                    $rulesSummary += "$($rule.'@odata.type' -replace '#microsoft.graph.',''): ID=$($rule.id)"
                }
            }
            "Policy: $($policy.displayName) | Scope: $($policy.scopeType) | Rules: $($rulesSummary -join ', ')"
        }
        return ($settingLines -join " || ")
    }
    catch {
        Write-Log "Fehler beim Lesen der PIM-Settings für Gruppe $GroupId: $_" -Level "WARN"
        return ""
    }
}

function Get-PimRoleAssignments {
    <#
    .SYNOPSIS Liest aktive PIM-Rollenzuweisungen (assignmentSchedules) für eine Gruppe.
    #>
    param([string]$GroupId)
    try {
        $uri         = "https://graph.microsoft.com/v1.0/identityGovernance/privilegedAccess/group/assignmentSchedules?`$filter=groupId eq '$GroupId'&`$expand=principal"
        $assignments = Get-GraphPagedResults -Uri $uri
        if (-not $assignments) { return "" }

        $lines = foreach ($a in $assignments) {
            $principal  = if ($a.principal.displayName) { $a.principal.displayName } else { $a.principalId }
            $role       = $a.accessId   # 'member' oder 'owner'
            $status     = $a.status
            $schedType  = if ($a.scheduleInfo.expiration.type) { $a.scheduleInfo.expiration.type } else { "permanent" }
            $expiry     = if ($a.scheduleInfo.expiration.endDateTime) { $a.scheduleInfo.expiration.endDateTime } else { "-" }
            "$principal | Rolle: $role | Status: $status | Ablauf: $schedType ($expiry)"
        }
        return ($lines -join " | ")
    }
    catch {
        Write-Log "Fehler beim Lesen der PIM-Assignments für Gruppe $GroupId: $_" -Level "WARN"
        return ""
    }
}

function Get-PimEligibleMembers {
    <#
    .SYNOPSIS Liest eligible (berechtigte) PIM-Mitglieder einer Gruppe.
    #>
    param([string]$GroupId)
    try {
        $uri      = "https://graph.microsoft.com/v1.0/identityGovernance/privilegedAccess/group/eligibilitySchedules?`$filter=groupId eq '$GroupId'&`$expand=principal"
        $eligible = Get-GraphPagedResults -Uri $uri
        if (-not $eligible) { return "" }

        $lines = foreach ($e in $eligible) {
            $principal  = if ($e.principal.displayName) { $e.principal.displayName } else { $e.principalId }
            $role       = $e.accessId
            $status     = $e.status
            $schedType  = if ($e.scheduleInfo.expiration.type) { $e.scheduleInfo.expiration.type } else { "permanent" }
            $expiry     = if ($e.scheduleInfo.expiration.endDateTime) { $e.scheduleInfo.expiration.endDateTime } else { "-" }
            "$principal | Rolle: $role | Status: $status | Ablauf: $schedType ($expiry)"
        }
        return ($lines -join " | ")
    }
    catch {
        Write-Log "Fehler beim Lesen der PIM-Eligible-Members für Gruppe $GroupId: $_" -Level "WARN"
        return ""
    }
}

#endregion

#region ── Hauptprogramm ─────────────────────────────────────────────────────

Write-Log "=== Entra ID Gruppen Export gestartet ===" -Level "OK"

# ── Modul prüfen ──────────────────────────────────────────────────────────────
if (-not (Get-Module -ListAvailable -Name Microsoft.Graph)) {
    Write-Log "Microsoft.Graph Modul nicht gefunden. Bitte installieren mit: Install-Module Microsoft.Graph" -Level "ERROR"
    exit 1
}

# ── Verbindung herstellen ─────────────────────────────────────────────────────
Write-Log "Verbinde mit Microsoft Graph..."
try {
    Connect-MgGraph -Scopes @(
        "Group.Read.All",
        "GroupMember.Read.All",
        "PrivilegedAccess.Read.AzureADGroup",
        "RoleManagementPolicy.Read.AzureADGroup"
    ) -ErrorAction Stop
    Write-Log "Verbindung erfolgreich hergestellt." -Level "OK"
}
catch {
    Write-Log "Fehler beim Verbinden mit Microsoft Graph: $_" -Level "ERROR"
    exit 1
}

# ── Alle Gruppen laden ────────────────────────────────────────────────────────
Write-Log "Lade alle Gruppen aus Entra ID..."
$groupUri  = "https://graph.microsoft.com/v1.0/groups?`$select=id,displayName,groupTypes,mailEnabled,securityEnabled,description&`$top=999"
$allGroups = Get-GraphPagedResults -Uri $groupUri
Write-Log "$($allGroups.Count) Gruppen gefunden." -Level "OK"

# ── Verarbeitung ──────────────────────────────────────────────────────────────
$exportData  = [System.Collections.Generic.List[PSCustomObject]]::new()
$totalGroups = $allGroups.Count
$counter     = 0

foreach ($group in $allGroups) {
    $counter++
    $pct = [math]::Round(($counter / $totalGroups) * 100)
    Write-Progress -Activity "Verarbeite Gruppen" -Status "$counter/$totalGroups – $($group.displayName)" -PercentComplete $pct

    Write-Log "[$counter/$totalGroups] Verarbeite: $($group.displayName)"

    # Grunddaten
    $groupName = $group.displayName
    $groupType = Get-GroupType -Group $group
    $members   = Get-GroupMembers -GroupId $group.id

    # PIM
    $pimEnabled      = Get-PimStatus       -GroupId $group.id
    $pimSettings     = ""
    $pimAssignments  = ""
    $pimEligible     = ""

    if ($pimEnabled) {
        Write-Log "  -> PIM for Groups ist aktiv. Lese PIM-Details..." -Level "OK"
        $pimSettings    = Get-PimSettings        -GroupId $group.id
        $pimAssignments = Get-PimRoleAssignments  -GroupId $group.id
        $pimEligible    = Get-PimEligibleMembers  -GroupId $group.id
    }

    $exportData.Add([PSCustomObject]@{
        Gruppenname             = $groupName
        Typ                     = $groupType
        Beschreibung            = $group.description
        Member                  = $members
        PIM_aktiviert           = if ($pimEnabled) { "Ja" } else { "Nein" }
        PIM_Einstellungen       = $pimSettings
        PIM_Rollenzuweisungen   = $pimAssignments
        PIM_Eligible_Members    = $pimEligible
    })
}

Write-Progress -Activity "Verarbeite Gruppen" -Completed

# ── CSV-Export ────────────────────────────────────────────────────────────────
Write-Log "Exportiere Daten nach: $OutputPath"
try {
    $exportData | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8 -Delimiter ";" -ErrorAction Stop
    Write-Log "Export erfolgreich abgeschlossen! Datei: $OutputPath" -Level "OK"
    Write-Log "Gesamt: $($exportData.Count) Gruppen exportiert." -Level "OK"
}
catch {
    Write-Log "Fehler beim Schreiben der CSV-Datei: $_" -Level "ERROR"
}

# ── Verbindung trennen ────────────────────────────────────────────────────────
Disconnect-MgGraph | Out-Null
Write-Log "Verbindung getrennt. Script beendet." -Level "OK"

#endregion
