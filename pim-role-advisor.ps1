<#
    .NAME
    PIM Role Advisor (AI-edition)

    .SYNOPSIS

    .NOTES
    
    .VERSION
    1.0
    
    .AUTHOR
    Morten Knudsen, Microsoft MVP - https://mortenknudsen.net

    .LICENSE
    Licensed under the MIT license.

    .PROJECTURI
    https://github.com/KnudsenMorten/PIM-Role-Advisor


    .WARRANTY
    Use at your own risk, no warranty given!
#>

# --- Connect to Microsoft Graph ---
if (-not (Get-MgContext)) {
    try {
         Connect-MgGraph -Scopes @(
            "PrivilegedAccess.Read.AzureAD",
            "RoleManagement.Read.Directory",
            "Directory.Read.All",
            "Group.Read.All"
         )
         Write-Host "✅ Connected to Microsoft Graph"
    } catch {
        throw "❌ Failed to connect to Microsoft Graph: $($_.Exception.Message)"
    }
}

# --- Connect to Azure (Az module) ---
if (-not (Get-AzContext)) {
    try {
        Connect-AzAccount -ErrorAction Stop
        Write-Host "✅ Connected to Azure"
    } catch {
        throw "❌ Failed to connect to Azure: $($_.Exception.Message)"
    }
}

    
# ================================================
# VARIABLES
# ================================================
$apiKey = "<Azure Open AI API Key>"
$endpoint   = "<endpoint name Azure Open AI>" #sample: "https://pim-role-advisor.openai.azure.com"
$deployment = "<deployment model name>"  # sample: "gpt-4o-mini"
$apiVersion = "2024-12-01-preview"
$uri = "$endpoint/openai/deployments/$deployment/chat/completions?api-version=$apiVersion"
$datapath = "c:\DATA"

# ================================================
# DATA PATHS
# ================================================
if (-not (Test-Path $datapath)) { New-Item -Path $datapath -ItemType Directory }

$entraRolesCsv     = "$datapath\ai-pim-role-advisor-EntraID-Roles.csv"
$pimGroupsCsv      = "$datapath\ai-pim-role-advisor-PIM-Groups-EntraID.csv"
$pimAzGroupsCsv    = "$datapath\ai-pim-role-advisor-PIM-Groups-AzRes.csv"
$entraEnrichedCsv  = "$datapath\ai-pim-role-advisor-EntraID-Enriched.csv"
$azEnrichedCsv     = "$datapath\ai-pim-role-advisor-AzRes-Enriched.csv"

# ================================================
# ENSURE POWERSHELL 7+
# ================================================

    if ($PSVersionTable.PSVersion.Major -lt 7) {
        throw "This script requires PowerShell 7 or later."
    }


function Is-FreshFile($path) {
    return (Test-Path $path) -and ((Get-Date) - (Get-Item $path).LastWriteTime).TotalMinutes -lt 60
}

#########################################################################################################
# Step 2: Getting list of Entra Role definitions
#########################################################################################################

Write-Progress -Activity "Loading Data" -Status "Step 1 of 5: Entra Role Definitions"
if (-not (Is-FreshFile $entraRolesCsv)) {
    $entraRoles = Get-MgBetaRoleManagementDirectoryRoleDefinition | Select-Object DisplayName, Id
    $entraRoles | Export-Csv -Path $entraRolesCsv -NoTypeInformation -Force -Delimiter ';'
    Write-Host "✅ Fetched Entra ID Roles..."
} else {
    Write-Host "✅ Loading Entra ID Roles from cache..."
    $entraRoles = Import-Csv -Path $entraRolesCsv -Delimiter ';'
}

$entraRoleMap = [System.Collections.Hashtable]::Synchronized(@{})
foreach ($role in $entraRoles) { $entraRoleMap[$role.Id] = $role.DisplayName }

#########################################################################################################
# Step 2: Getting list of PIM-Entra groups
#########################################################################################################

Write-Progress -Activity "Loading Data" -Status "Step 2 of 5: PIM-Entra Groups"
if (-not (Is-FreshFile $pimGroupsCsv)) {
    $pimGroups = Get-MgGroup -Filter "startswith(displayName,'PIM-Entra')" -all | Select-object Id, DisplayName
    $pimGroups | Export-Csv -Path $pimGroupsCsv -NoTypeInformation -Force -Delimiter ';'
    Write-Host "✅ Fetched PIM Entra ID Groups"
} else {
    Write-Host "✅ Loading PIM Entra ID Groups from cache..."
    $pimGroups = Import-Csv -Path $pimGroupsCsv -Delimiter ';'
}

#########################################################################################################
# Step 3: Getting list of PIM-AzRes groups
#########################################################################################################

Write-Progress -Activity "Loading Data" -Status "Step 3 of 5: PIM-AzRes Groups"
if (-not (Is-FreshFile $pimAzGroupsCsv)) {
    $pimAzGroups = Get-MgGroup -Filter "startswith(displayName,'PIM-AzRes-')" -all | Select-object Id, DisplayName
    $pimAzGroups | Export-Csv -Path $pimAzGroupsCsv -NoTypeInformation -Force -Delimiter ';'
    Write-Host "✅ Fetched PIM Azure Groups"
} else {
    Write-Host "✅ Loading PIM Azure Groups from cache..."
    $pimAzGroups = Import-Csv -Path $pimAzGroupsCsv -Delimiter ';'
}

#########################################################################################################
# Step 4: Enriching Enta ID Role Assignments
#########################################################################################################

Write-Progress -Activity "Enriching Data" -Status "Step 4 of 5: Enriching Entra ID Role Assignments"
if (-not (Is-FreshFile $entraEnrichedCsv)) {
    # Set maximum parallel threads
    $maxThreads = 10
    $totalGroups = $pimGroups.Count

    # Create runspace pool
    $runspacePool = [runspacefactory]::CreateRunspacePool(1, $maxThreads)
    $runspacePool.ThreadOptions = "ReuseThread"
    $runspacePool.Open()

    # Prepare thread-safe collection
    Add-Type -AssemblyName System.Collections
    $pimEntraGroupEnriched = [System.Collections.Concurrent.ConcurrentBag[object]]::new()
    $jobs = [System.Collections.Generic.List[object]]::new()

    # Start runspaces for each group
    foreach ($group in $pimGroups) {
        $ps = [powershell]::Create()
        $ps.RunspacePool = $runspacePool

        [void]$ps.AddScript({
            param ($group, $entraRoleMap)

            $groupId = $group.Id
            $groupName = $group.DisplayName

            try {
                $active = @(
                    (Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/beta/roleManagement/directory/roleAssignments?`$filter=principalId eq '$groupId'" -Method GET).value |
                    ForEach-Object {
                        $rid = $_.roleDefinitionId
                        $rname = if ($entraRoleMap.ContainsKey($rid)) { $entraRoleMap[$rid] } else { "Unknown Role" }
                        "$rname [$rid] (Active)"
                    }
                ) -join ", "
                if (-not $active) { $active = "None" }
            } catch {
                $active = "Error"
            }

            try {
                $eligible = @(
                    (Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/beta/roleManagement/directory/roleEligibilityScheduleInstances?`$filter=principalId eq '$groupId'" -Method GET).value |
                    ForEach-Object {
                        $rid = $_.roleDefinitionId
                        $rname = if ($entraRoleMap.ContainsKey($rid)) { $entraRoleMap[$rid] } else { "Unknown Role" }
                        "$rname [$rid] (Eligible)"
                    }
                ) -join ", "
                if (-not $eligible) { $eligible = "None" }
            } catch {
                $eligible = "Error"
            }

            return [PSCustomObject]@{
                DisplayName         = $groupName
                Id                  = $groupId
                ActiveAssignments   = $active
                EligibleAssignments = $eligible
            }

        }).AddArgument($group).AddArgument($entraRoleMap) | Out-Null

        $job = $ps.BeginInvoke()
        $jobInfo = [PSCustomObject]@{
            Pipe      = $ps
            Handle    = $job
            GroupName = $group.DisplayName
            Processed = $false
        }

        [void]$jobs.Add($jobInfo)
    }

    # Wait for all jobs with progress bar
    $completed = 0
    while ($completed -lt $jobs.Count) {
        for ($i = 0; $i -lt $jobs.Count; $i++) {
            $job = $jobs[$i]
            if ($job.Handle.IsCompleted -and -not $job.Processed) {
                $result = $job.Pipe.EndInvoke($job.Handle)
                foreach ($entry in $result) {
                    $pimEntraGroupEnriched.Add($entry)
                }
                $job.Pipe.Dispose()
                $job.Processed = $true
                $completed++

                $percent = [math]::Min([math]::Round(($completed / $totalGroups) * 100), 100)
                Write-Progress -Activity "Enriching PIM-Entra Groups" -Status "Processing group $completed of $totalGroups" -PercentComplete $percent
            }
        }
        Start-Sleep -Milliseconds 100
    }
    
    $pimEntraGroupEnriched | Export-Csv -Path $entraEnrichedCsv -NoTypeInformation -Force -Delimiter ';'

    Write-Host "✅ Enriched Entra ID Assignments"

    Write-Progress -Activity "Enriching PIM-Entra Groups" -Completed

    # Cleanup runspace pool
    $runspacePool.Close()
    $runspacePool.Dispose()

} else {
    Write-Host "✅ Loading Enriched Entra ID Assignments from cache..."
    $pimEntraGroupEnriched = Import-Csv -Path $entraEnrichedCsv -Delimiter ';'
}

#########################################################################################################
# Step 5: Enriching Azure Group Assignments
#########################################################################################################

Write-Progress -Activity "Enriching Data" -Status "Step 5 of 5: Enriching Azure Group Assignments"

if (-not (Is-FreshFile $azEnrichedCsv)) {
    $maxThreads = 10
    $totalGroups = $pimAzGroups.Count

    $runspacePool = [runspacefactory]::CreateRunspacePool(1, $maxThreads)
    $runspacePool.ThreadOptions = "ReuseThread"
    $runspacePool.Open()

    $pimAzGroupEnriched = [System.Collections.Concurrent.ConcurrentBag[object]]::new()
    $jobs = [System.Collections.Generic.List[object]]::new()

    foreach ($group in $pimAzGroups) {
        $ps = [powershell]::Create()
        $ps.RunspacePool = $runspacePool

        [void]$ps.AddScript({
            param ($group)

            Import-Module Az.Resources -Force

            $groupId = $group.Id
            $groupName = $group.DisplayName

            try {
                $assignments = Get-AzRoleAssignment -ObjectId $groupId -ErrorAction Stop

                $mapped = $assignments | ForEach-Object {
                    "$($_.RoleDefinitionName) [$($_.RoleDefinitionId)] on $($_.Scope)"
                }

                if (-not $mapped) { $mapped = @("None") }
            } catch {
                $mapped = @("Error: $($_.Exception.Message)")
            }

            return [PSCustomObject]@{
                DisplayName       = $groupName
                Id                = $groupId
                AzureRoleMappings = $mapped -join ", "
            }
        }).AddArgument($group) | Out-Null



        $job = $ps.BeginInvoke()
        $jobInfo = [PSCustomObject]@{
            Pipe      = $ps
            Handle    = $job
            GroupName = $group.DisplayName
            Processed = $false
        }

        [void]$jobs.Add($jobInfo)
    }

    # Progress tracking
    $completed = 0
    $progressId = 42

    while ($completed -lt $jobs.Count) {
        $jobs | ForEach-Object {
            if (-not $_.Processed -and $_.Handle.IsCompleted) {
                $result = $_.Pipe.EndInvoke($_.Handle)
                foreach ($entry in $result) {
                    $pimAzGroupEnriched.Add($entry)
                }                
                $_.Pipe.Dispose()
                $_.Processed = $true
                $completed++
            }
        }

        $percent = [math]::Min([math]::Round(($completed / $totalGroups) * 100), 100)
        Write-Progress -Id $progressId -Activity "Enriching Azure Group Assignments" -Status "Processed $completed of $totalGroups" -PercentComplete $percent
        Start-Sleep -Milliseconds 150
    }

    Write-Progress -Id $progressId -Activity "Enriching Azure Group Assignments" -Completed
    $runspacePool.Close()
    $runspacePool.Dispose()

    $pimAzGroupEnriched | Export-Csv -Path $azEnrichedCsv -NoTypeInformation -Force -Delimiter ';'
    Write-Host "✅ Enriched Azure Anrichments"

} else {
    Write-Host "✅ Loading Enriched Azure Assignments data from cache..."
    $pimAzGroupEnriched = Import-Csv -Path $azEnrichedCsv -Delimiter ';'
}
Write-Progress -Activity "Enriching Data" -Status "Step 5 of 5: Enriching Azure Group Assignments" -Completed


#########################################################################################################
# BUILD AI PROMPT CONTEXT
#########################################################################################################

$entraRoleText = ($entraRoles | ForEach-Object { "$($_.DisplayName) [$($_.Id)]" }) -join "`n"
$azureRoleText = (Get-AzRoleDefinition | ForEach-Object { "$($_.Name) [$($_.Id)]" }) -join "`n"
$pimGroupEnrichedText = ($pimEntraGroupEnriched | ForEach-Object {
    @"
$($_.DisplayName) [$($_.Id)]
  Active Assignments:   $($_.ActiveAssignments)
  Eligible Assignments: $($_.EligibleAssignments)
"@
}) -join "`n"
$pimAzGroupText = ($pimAzGroupEnriched | ForEach-Object {
    @"
$($_.DisplayName) [$($_.Id)]
  Azure Role Assignments: $($_.AzureRoleMappings)
"@
}) -join "`n"

#########################################################################################################
# BUILD AND DISPLAY GUI
#########################################################################################################

Add-Type -AssemblyName System.Windows.Forms

while ($true) {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "PIM Role Advisor (AI-edition) | Created by Microsoft MVP, Morten Knudsen"
    $form.Font = New-Object System.Drawing.Font("Segoe UI", 14, [System.Drawing.FontStyle]::Bold)
    $form.Width = 600
    $form.StartPosition = "CenterScreen"

    $label = New-Object System.Windows.Forms.Label
    $label.Text = "Ask your question below (e.g. 'create an Azure App', 'read secret from Key Vault')"
    $label.AutoSize = $true
    $label.MaximumSize = New-Object System.Drawing.Size(560, 0)
    $label.Font = New-Object System.Drawing.Font("Segoe UI", 16, [System.Drawing.FontStyle]::Bold)
    $label.Location = New-Object System.Drawing.Point(10, 10)
    $form.Controls.Add($label)

    $textbox = New-Object System.Windows.Forms.TextBox
    $textbox.Multiline = $true
    $textbox.Width = 560
    $textbox.Height = 150
    $textbox.Location = New-Object System.Drawing.Point(10, ($label.Top + $label.Height + 10))
    $textbox.Font = New-Object System.Drawing.Font("Segoe UI", 16, [System.Drawing.FontStyle]::Bold)
    $form.Controls.Add($textbox)

    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Text = "Ask"
    $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $okButton.Location = New-Object System.Drawing.Point(400, ($textbox.Top + $textbox.Height + 10))
    $okButton.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
    $form.AcceptButton = $okButton
    $form.Controls.Add($okButton)

    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Text = "Cancel"
    $cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $cancelButton.Location = New-Object System.Drawing.Point(480, ($textbox.Top + $textbox.Height + 10))
    $cancelButton.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
    $form.CancelButton = $cancelButton
    $form.Controls.Add($cancelButton)

    $form.Height = ($okButton.Top + $okButton.Height + 70)
    $dialogResult = $form.ShowDialog()

    if ($dialogResult -ne [System.Windows.Forms.DialogResult]::OK) { break }
    $customQuestion = $textbox.Text.Trim()
    if (-not $customQuestion) { continue }

#########################################################################################################
# Send Prompt to Azure Open AI | Stream result to Screen, when output is available
#########################################################################################################

    $userPrompt = @"
You are a role advisor AI. Based on the following context, answer the user's question.

Entra ID Roles:
$entraRoleText

Azure RBAC Roles:
$azureRoleText

PIM Entra Groups Role Assignments:
$pimGroupEnrichedText

PIM Azure Resource Groups Role Assignments:
$pimAzGroupText

User Question:
$customQuestion

Return format:
- Role Name => Recommended PIM Group(s)
Prioritize the least privileged group that fulfills the requirements.
"@

    # OpenAI Request
    Write-Host "`n[AI RESPONSE]`n" -ForegroundColor Cyan
    try {
        $body = @{
            model = $deployment
            stream = $true
            temperature = 0.7
            top_p = 1.0
            max_tokens = 4096
            messages = @(
                @{ role = "system"; content = "You are a helpful assistant who advises on access governance and role-to-group assignments." },
                @{ role = "user"; content = $userPrompt }
            )
        } | ConvertTo-Json -Depth 10 -Compress

        $handler = [System.Net.Http.HttpClientHandler]::new()
        $client = [System.Net.Http.HttpClient]::new($handler)
        $request = [System.Net.Http.HttpRequestMessage]::new([System.Net.Http.HttpMethod]::Post, $uri)

        $request.Headers.Add("api-key", $apiKey)
        $request.Headers.Add("Accept", "text/event-stream")
        $request.Content = [System.Net.Http.StringContent]::new($body, [System.Text.Encoding]::UTF8, "application/json")

        $response = $client.SendAsync($request, [System.Net.Http.HttpCompletionOption]::ResponseHeadersRead).Result
        $stream = $response.Content.ReadAsStreamAsync().Result
        $reader = [System.IO.StreamReader]::new($stream)

        while (-not $reader.EndOfStream) {
            $line = $reader.ReadLine()
            if ($line -and $line.StartsWith("data: ")) {
                $json = $line.Substring(6)
                if ($json -eq "[DONE]") { break }
                try {
                    $obj = $json | ConvertFrom-Json
                    $text = $obj.choices[0].delta.content
                    if ($text) { Write-Host -NoNewline $text }
                } catch {
                    Write-Warning "Failed to parse chunk: $json"
                }
            }
        }
        $reader.Close()
        $client.Dispose()
    } catch {
        Write-Error "OpenAI request failed: $($_.Exception.Message)"
        if ($reader) { $reader.Close() }
        if ($client) { $client.Dispose() }
    }
}
