# PowerShell Script: Hyper-V VM Replication Enabler with GUI for Cluster

# Load required assemblies for GUI
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName System.Windows.Forms

# Import required modules
Import-Module FailoverClusters
Import-Module Hyper-V

# Create the main form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Enable Hyper-V VM Replication (Cluster)"
$form.Size = New-Object System.Drawing.Size(800, 720)
$form.StartPosition = "CenterScreen"

# Label for Cluster Name
$labelClusterName = New-Object System.Windows.Forms.Label
$labelClusterName.Text = "Cluster Name:"
$labelClusterName.Location = New-Object System.Drawing.Point(10, 10)
$labelClusterName.Size = New-Object System.Drawing.Size(150, 20)
$form.Controls.Add($labelClusterName)

# TextBox for Cluster Name
$textBoxClusterName = New-Object System.Windows.Forms.TextBox
$textBoxClusterName.Location = New-Object System.Drawing.Point(170, 10)
$textBoxClusterName.Size = New-Object System.Drawing.Size(400, 20)
$form.Controls.Add($textBoxClusterName)

# Button to Load VMs
$buttonLoadVMs = New-Object System.Windows.Forms.Button
$buttonLoadVMs.Text = "Load VMs"
$buttonLoadVMs.Location = New-Object System.Drawing.Point(10, 40)
$buttonLoadVMs.Size = New-Object System.Drawing.Size(150, 30)
$form.Controls.Add($buttonLoadVMs)

# Label for VM selection
$labelVMs = New-Object System.Windows.Forms.Label
$labelVMs.Text = "Select VMs:"
$labelVMs.Location = New-Object System.Drawing.Point(10, 80)
$labelVMs.Size = New-Object System.Drawing.Size(200, 20)
$form.Controls.Add($labelVMs)

# Checkbox for Select All VMs
$checkBoxSelectAll = New-Object System.Windows.Forms.CheckBox
$checkBoxSelectAll.Text = "Select All VMs"
$checkBoxSelectAll.Location = New-Object System.Drawing.Point(10, 100)
$checkBoxSelectAll.Size = New-Object System.Drawing.Size(150, 20)
$form.Controls.Add($checkBoxSelectAll)

# CheckedListBox for VMs
$checkedListBoxVMs = New-Object System.Windows.Forms.CheckedListBox
$checkedListBoxVMs.Location = New-Object System.Drawing.Point(10, 130)
$checkedListBoxVMs.Size = New-Object System.Drawing.Size(760, 150)
$form.Controls.Add($checkedListBoxVMs)

# Checkbox for Filtering VMs with replication
$checkBoxFilter = New-Object System.Windows.Forms.CheckBox
$checkBoxFilter.Text = "Hide VMs with replication enabled"
$checkBoxFilter.Location = New-Object System.Drawing.Point(10, 290)
$checkBoxFilter.Size = New-Object System.Drawing.Size(200, 20)
$checkBoxFilter.Checked = $true
$form.Controls.Add($checkBoxFilter)

# Label for Replica Server
$labelReplicaServer = New-Object System.Windows.Forms.Label
$labelReplicaServer.Text = "Replica Server Name:"
$labelReplicaServer.Location = New-Object System.Drawing.Point(10, 320)
$labelReplicaServer.Size = New-Object System.Drawing.Size(150, 20)
$form.Controls.Add($labelReplicaServer)

# TextBox for Replica Server
$textBoxReplicaServer = New-Object System.Windows.Forms.TextBox
$textBoxReplicaServer.Location = New-Object System.Drawing.Point(170, 320)
$textBoxReplicaServer.Size = New-Object System.Drawing.Size(400, 20)
$form.Controls.Add($textBoxReplicaServer)

# Label for Replica Port
$labelReplicaPort = New-Object System.Windows.Forms.Label
$labelReplicaPort.Text = "Replica Server Port:"
$labelReplicaPort.Location = New-Object System.Drawing.Point(10, 350)
$labelReplicaPort.Size = New-Object System.Drawing.Size(150, 20)
$form.Controls.Add($labelReplicaPort)

# TextBox for Replica Port (default 80 for HTTP, 443 for HTTPS)
$textBoxReplicaPort = New-Object System.Windows.Forms.TextBox
$textBoxReplicaPort.Location = New-Object System.Drawing.Point(170, 350)
$textBoxReplicaPort.Size = New-Object System.Drawing.Size(400, 20)
$textBoxReplicaPort.Text = "80"
$form.Controls.Add($textBoxReplicaPort)

# Label for Authentication Type
$labelAuthType = New-Object System.Windows.Forms.Label
$labelAuthType.Text = "Authentication Type:"
$labelAuthType.Location = New-Object System.Drawing.Point(10, 380)
$labelAuthType.Size = New-Object System.Drawing.Size(150, 20)
$form.Controls.Add($labelAuthType)

# ComboBox for Authentication Type
$comboBoxAuthType = New-Object System.Windows.Forms.ComboBox
$comboBoxAuthType.Location = New-Object System.Drawing.Point(170, 380)
$comboBoxAuthType.Size = New-Object System.Drawing.Size(400, 20)
$comboBoxAuthType.Items.AddRange(@("Kerberos", "Certificate"))
$comboBoxAuthType.SelectedIndex = 0
$form.Controls.Add($comboBoxAuthType)

# Label for Replication Frequency (in seconds)
$labelFrequency = New-Object System.Windows.Forms.Label
$labelFrequency.Text = "Replication Frequency (sec):"
$labelFrequency.Location = New-Object System.Drawing.Point(10, 410)
$labelFrequency.Size = New-Object System.Drawing.Size(150, 20)
$form.Controls.Add($labelFrequency)

# ComboBox for Frequency
$comboBoxFrequency = New-Object System.Windows.Forms.ComboBox
$comboBoxFrequency.Location = New-Object System.Drawing.Point(170, 410)
$comboBoxFrequency.Size = New-Object System.Drawing.Size(400, 20)
$comboBoxFrequency.Items.AddRange(@("30", "300", "900"))  # 30s for Hyper-V 2016+, 5min, 15min
$comboBoxFrequency.SelectedIndex = 1  # Default to 300 (5 min)
$form.Controls.Add($comboBoxFrequency)

# Label for Replica Base Path
$labelBasePath = New-Object System.Windows.Forms.Label
$labelBasePath.Text = "Base Path for Move:"
$labelBasePath.Location = New-Object System.Drawing.Point(10, 440)
$labelBasePath.Size = New-Object System.Drawing.Size(150, 20)
$form.Controls.Add($labelBasePath)

# TextBox for Replica Base Path
$textBoxBasePath = New-Object System.Windows.Forms.TextBox
$textBoxBasePath.Location = New-Object System.Drawing.Point(170, 440)
$textBoxBasePath.Size = New-Object System.Drawing.Size(400, 20)
$textBoxBasePath.Text = "C:\ClusterStorage\Volume1\Replication"
$form.Controls.Add($textBoxBasePath)

# Checkbox for Compression
$checkBoxCompression = New-Object System.Windows.Forms.CheckBox
$checkBoxCompression.Text = "Enable Compression"
$checkBoxCompression.Location = New-Object System.Drawing.Point(10, 470)
$checkBoxCompression.Size = New-Object System.Drawing.Size(150, 20)
$checkBoxCompression.Checked = $true
$form.Controls.Add($checkBoxCompression)

# Button to Enable Replication
$buttonEnable = New-Object System.Windows.Forms.Button
$buttonEnable.Text = "Enable Replication"
$buttonEnable.Location = New-Object System.Drawing.Point(10, 500)
$buttonEnable.Size = New-Object System.Drawing.Size(150, 30)
$form.Controls.Add($buttonEnable)

# Button to Force Replication Now
$buttonForceReplication = New-Object System.Windows.Forms.Button
$buttonForceReplication.Text = "Force Replication Now"
$buttonForceReplication.Location = New-Object System.Drawing.Point(170, 500)
$buttonForceReplication.Size = New-Object System.Drawing.Size(150, 30)
$form.Controls.Add($buttonForceReplication)

# Button to Move VMs
$buttonMove = New-Object System.Windows.Forms.Button
$buttonMove.Text = "Move VMs"
$buttonMove.Location = New-Object System.Drawing.Point(10, 540)
$buttonMove.Size = New-Object System.Drawing.Size(150, 30)
$form.Controls.Add($buttonMove)

# Button to Generate Replication Health Report
$buttonHealthReport = New-Object System.Windows.Forms.Button
$buttonHealthReport.Text = "Replication Health Report"
$buttonHealthReport.Location = New-Object System.Drawing.Point(170, 540)
$buttonHealthReport.Size = New-Object System.Drawing.Size(150, 30)
$form.Controls.Add($buttonHealthReport)

# Status Label
$labelStatus = New-Object System.Windows.Forms.Label
$labelStatus.Text = ""
$labelStatus.Location = New-Object System.Drawing.Point(10, 580)
$labelStatus.Size = New-Object System.Drawing.Size(760, 50)
$labelStatus.ForeColor = "Blue"
$form.Controls.Add($labelStatus)

# Global variables to store VMs, their hosts, and replica servers
$global:vmList = @()
$global:vmHosts = @{}
$global:vmReplicaServers = @{}

# Function to populate VMs based on filter
function PopulateVMs {
    $checkedListBoxVMs.Items.Clear()
    foreach ($vm in $global:vmList) {
        $rep = Get-VMReplication -VMName $vm.Name -ComputerName $vm.ComputerName -ErrorAction SilentlyContinue
        $hasRep = $null -ne $rep
        if (-not $checkBoxFilter.Checked -or -not $hasRep) {
            # Get the path of the first VHD/VHDX
            $hardDiskPath = Invoke-Command -ComputerName $vm.ComputerName -ScriptBlock {
                $hdd = Get-VMHardDiskDrive -VMName $using:vm.Name -ErrorAction SilentlyContinue | Select-Object -First 1
                if ($hdd) { $hdd.Path } else { "No Hard Disk Found" }
            }
            $displayText = "$($vm.Name) - $hardDiskPath"
            $checkedListBoxVMs.Items.Add($displayText)
        }
    }
    # Update Select All checkbox state
    $checkBoxSelectAll.Checked = $false
}

# Event Handler for Select All Checkbox
$checkBoxSelectAll.Add_CheckedChanged({
    for ($i = 0; $i -lt $checkedListBoxVMs.Items.Count; $i++) {
        $checkedListBoxVMs.SetItemChecked($i, $checkBoxSelectAll.Checked)
    }
})

# Event Handler for Load VMs Button
$buttonLoadVMs.Add_Click({
    $clusterName = $textBoxClusterName.Text
    if ([string]::IsNullOrEmpty($clusterName)) {
        $labelStatus.Text = "Cluster Name is required."
        return
    }

    try {
        # Get VMs from cluster using Get-ClusterResource
        $global:vmList = Get-ClusterResource -Cluster $clusterName | Where-Object ResourceType -eq "Virtual Machine" | Get-VM | Sort-Object Name
        if ($global:vmList.Count -eq 0) {
            $labelStatus.Text = "No VMs found in cluster."
            return
        }

        # Store VM host and replica server information
        $global:vmHosts = @{}
        $global:vmReplicaServers = @{}
        foreach ($vm in $global:vmList) {
            $global:vmHosts[$vm.Name] = $vm.ComputerName
            $rep = Get-VMReplication -VMName $vm.Name -ComputerName $vm.ComputerName -ErrorAction SilentlyContinue
            if ($rep) {
                $global:vmReplicaServers[$vm.Name] = $rep.ReplicaServerName
            }
        }

        # Populate VMs
        PopulateVMs
        $labelStatus.Text = "VMs loaded successfully."
    } catch {
        $labelStatus.Text = "Error loading VMs: $($_.Exception.Message)"
    }
})

# Event Handler for Filter Checkbox
$checkBoxFilter.Add_CheckedChanged({ PopulateVMs })

# Function to extract VM name from display text
function Get-VMNameFromDisplay {
    param($displayText)
    if ($displayText -match "^(.+?)\s*-\s*(.+)") {
        return $matches[1].Trim()
    } else {
        return $displayText
    }
}

# Event Handler for Enable Replication Button
$buttonEnable.Add_Click({
    $selectedDisplays = $checkedListBoxVMs.CheckedItems
    if ($selectedDisplays.Count -eq 0) {
        $labelStatus.Text = "No VMs selected."
        return
    }
    if ([string]::IsNullOrEmpty($textBoxReplicaServer.Text)) {
        $labelStatus.Text = "Replica Server Name is required."
        return
    }

    $selectedVMs = $selectedDisplays | ForEach-Object { Get-VMNameFromDisplay $_ }
    $replicaServer = $textBoxReplicaServer.Text
    $replicaPort = [int]$textBoxReplicaPort.Text
    $authType = $comboBoxAuthType.SelectedItem
    $frequency = [int]$comboBoxFrequency.SelectedItem
    $compression = $checkBoxCompression.Checked
    $totalVMs = $selectedVMs.Count
    $currentVM = 0

    try {
        foreach ($name in $selectedVMs) {
            $currentVM++
            $percentComplete = [math]::Round(($currentVM / $totalVMs) * 100, 2)
            $labelStatus.Text = "Processing VM $currentVM/$totalVMs ($percentComplete% complete): Enabling replication for $name"
            [System.Windows.Forms.Application]::DoEvents()

            $hostName = $global:vmHosts[$name]
            Enable-VMReplication -VMName $name -ComputerName $hostName -ReplicaServerName $replicaServer -ReplicaServerPort $replicaPort -AuthenticationType $authType -CompressionEnabled $compression -ReplicationFrequencySec $frequency
            Start-VMInitialReplication -VMName $name -ComputerName $hostName  # Optionally start initial replication
            $global:vmReplicaServers[$name] = $replicaServer  # Update replica server mapping
        }
        $labelStatus.Text = "Replication enabled for $totalVMs VMs (100% complete)."
    } catch {
        $labelStatus.Text = "Error: $($_.Exception.Message)"
    }
})

# Event Handler for Move VMs Button
$buttonMove.Add_Click({
    $selectedDisplays = $checkedListBoxVMs.CheckedItems
    if ($selectedDisplays.Count -eq 0) {
        $labelStatus.Text = "No VMs selected."
        return
    }
    if ([string]::IsNullOrEmpty($textBoxBasePath.Text)) {
        $labelStatus.Text = "Base Path is required."
        return
    }

    $selectedVMs = $selectedDisplays | ForEach-Object { Get-VMNameFromDisplay $_ }
    $basePath = $textBoxBasePath.Text
    $totalVMs = $selectedVMs.Count
    $currentVM = 0

    try {
        foreach ($name in $selectedVMs) {
            $currentVM++
            $percentComplete = [math]::Round(($currentVM / $totalVMs) * 100, 2)
            $labelStatus.Text = "Processing VM $currentVM/$totalVMs ($percentComplete% complete): Moving $name"
            [System.Windows.Forms.Application]::DoEvents()

            $hostName = $global:vmHosts[$name]
            $newFolder = Join-Path -Path $basePath -ChildPath $name

            # Create folder on the VM's host
            Invoke-Command -ComputerName $hostName -ScriptBlock {
                New-Item -Path $using:newFolder -ItemType Directory -Force
            }

            # Move storage to the named subfolder
            Invoke-Command -ComputerName $hostName -ScriptBlock {
                Move-VMStorage -VMName $using:name -DestinationStoragePath $using:newFolder
            }
        }
        $labelStatus.Text = "Moved $totalVMs VMs to their named subfolders (100% complete)."
    } catch {
        $labelStatus.Text = "Error: $($_.Exception.Message)"
    }
})

# Event Handler for Force Replication Now Button
$buttonForceReplication.Add_Click({
    $selectedDisplays = $checkedListBoxVMs.CheckedItems
    if ($selectedDisplays.Count -eq 0) {
        $labelStatus.Text = "No VMs selected."
        return
    }

    $selectedVMs = $selectedDisplays | ForEach-Object { Get-VMNameFromDisplay $_ }
    $totalVMs = $selectedVMs.Count
    $currentVM = 0

    try {
        foreach ($name in $selectedVMs) {
            $currentVM++
            $percentComplete = [math]::Round(($currentVM / $totalVMs) * 100, 2)
            $labelStatus.Text = "Processing VM $currentVM/$totalVMs ($percentComplete% complete): Forcing replication for $name"
            [System.Windows.Forms.Application]::DoEvents()

            $hostName = $global:vmHosts[$name]
            $replicaServer = $global:vmReplicaServers[$name]
            if (-not $replicaServer) {
                $labelStatus.Text = "Replication not configured for VM $name ($percentComplete% complete)."
                continue
            }

            # Check if replication is enabled
            $rep = Get-VMReplication -VMName $name -ComputerName $hostName -ErrorAction SilentlyContinue
            if (-not $rep) {
                $labelStatus.Text = "Replication not enabled for VM $name ($percentComplete% complete)."
                continue
            }

            # Force replication
            Invoke-Command -ComputerName $hostName -ScriptBlock {
                Start-VMInitialReplication -VMName $using:name -Force
            }
        }
        $labelStatus.Text = "Forced replication started for $totalVMs VMs (100% complete)."
    } catch {
        $labelStatus.Text = "Error: $($_.Exception.Message)"
    }
})

# Event Handler for Replication Health Report Button
$buttonHealthReport.Add_Click({
    $selectedDisplays = $checkedListBoxVMs.CheckedItems
    if ($selectedDisplays.Count -eq 0) {
        $labelStatus.Text = "No VMs selected."
        return
    }

    # Prompt for save location using SaveFileDialog
    $saveDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveDialog.Filter = "HTML files (*.html)|*.html|All files (*.*)|*.*"
    $saveDialog.Title = "Save Replication Health Report"
    $saveDialog.FileName = "ReplicationHealthReport_$(Get-Date -Format 'yyyyMMdd_HHmmss').html"
    if ($saveDialog.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
        $labelStatus.Text = "Report generation cancelled."
        return
    }
    $reportPath = $saveDialog.FileName

    $selectedVMs = $selectedDisplays | ForEach-Object { Get-VMNameFromDisplay $_ }
    $totalVMs = $selectedVMs.Count
    $currentVM = 0
    $reportData = @()

    try {
        foreach ($name in $selectedVMs) {
            $currentVM++
            $percentComplete = [math]::Round(($currentVM / $totalVMs) * 100, 2)
            $labelStatus.Text = "Processing VM $currentVM/$totalVMs ($percentComplete% complete): Generating report for $name"
            [System.Windows.Forms.Application]::DoEvents()

            $hostName = $global:vmHosts[$name]
            $replicaServer = $global:vmReplicaServers[$name]
            if (-not $replicaServer) {
                $labelStatus.Text = "Replication not configured for VM $name ($percentComplete% complete)."
                continue
            }

            # Get replication health data
            $rep = Get-VMReplication -VMName $name -ComputerName $hostName -ErrorAction SilentlyContinue
            $measure = Measure-VMReplication -VMName $name -ComputerName $hostName -ErrorAction SilentlyContinue
            if ($rep -and $measure) {
                $reportData += [PSCustomObject]@{
                    VMName            = $name
                    HostName          = $hostName
                    ReplicaServer     = $rep.ReplicaServerName
                    ReplicationState  = $rep.State
                    ReplicationHealth = $rep.Health
                    LastReplicationTime = $measure.LReplTime
                    AvgReplicationSizeMB = $measure.AvgReplSize
                }
            } else {
                $labelStatus.Text = "Replication data not available for VM $name ($percentComplete% complete)."
                continue
            }
        }

        if ($reportData.Count -eq 0) {
            $labelStatus.Text = "No replication data available for selected VMs."
            return
        }

        # Generate HTML file
        $htmlHead = "<style>table {border-collapse: collapse; width: 100%;} th, td {border: 1px solid black; padding: 8px; text-align: left;} th {background-color: #f2f2f2;}</style>"
        $htmlBody = $reportData | ConvertTo-Html -Head $htmlHead -Body "<h2>Hyper-V Replication Health Report</h2><p>Generated on $(Get-Date)</p>" -PostContent "<p>Report includes $($reportData.Count) VMs.</p>"

        # Write to file
        $htmlBody | Out-File -FilePath $reportPath -Encoding UTF8

        # Open the report
        Start-Process $reportPath

        $labelStatus.Text = "Health report for $totalVMs VMs generated and opened at $reportPath (100% complete)."
    } catch {
        $labelStatus.Text = "Error generating report: $($_.Exception.Message)"
    }
})

# Show the form
$form.ShowDialog() | Out-Null
