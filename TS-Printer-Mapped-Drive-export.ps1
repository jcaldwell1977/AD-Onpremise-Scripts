# Collect-DriveMappingsAndPrinters.ps1
# PowerShell script to collect or import logged-in user's drive mappings and printers for SCCM
# Exports to or imports from a JSON file at a network location

param (
    [Parameter(Mandatory=$true)]
    [ValidateSet("Export","Import")]
    [string]$Mode,
    [Parameter(Mandatory=$true)]
    [string]$LogFilePath
)

# Define variables
$NetworkPath = "\\Server\Share\SCCM_Inventory" # Replace with your network share path
$OutputFileName = "Inventory_$($env:COMPUTERNAME)_$($env:USERNAME)_$(Get-Date -Format 'yyyyMMdd_HHmmss').json"
$OutputFilePath = Join-Path -Path $NetworkPath -ChildPath $OutputFileName
$User = $env:USERNAME
$Computer = $env:COMPUTERNAME
$DateTime = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

# Function to write logs
function Write-Log {
    param($Message)
    $LogMessage = "$DateTime - Computer: $Computer - User: $User - $Message"
    Add-Content -Path $LogFilePath -Value $LogMessage
}

try {
    Write-Log "Starting $Mode process for drive mappings and printers."

    # Ensure log file directory exists
    $LogDir = Split-Path -Path $LogFilePath -Parent
    if (-not (Test-Path $LogDir)) {
        New-Item -ItemType Directory -Path $LogDir -Force | Out-Null
        Write-Log "Created log directory at $LogDir."
    }

    if ($Mode -eq "Export") {
        # Ensure network path is accessible
        if (-not (Test-Path $NetworkPath)) {
            Write-Log "Network path $NetworkPath is inaccessible."
            exit 1
        }

        # Initialize output object
        $Inventory = @{
            ComputerName = $Computer
            UserName = $User
            Timestamp = $DateTime
            DriveMappings = @()
            Printers = @()
        }

        # Collect mapped network drives
        Write-Log "Collecting mapped network drives."
        $DriveMappings = Get-WmiObject -Class Win32_MappedLogicalDisk | ForEach-Object {
            @{
                DriveLetter = $_.DeviceID
                Path = $_.ProviderName
                Name = $_.Name
            }
        }
        if ($DriveMappings) {
            $Inventory.DriveMappings = $DriveMappings
            foreach ($Drive in $DriveMappings) {
                Write-Log "Mapped Drive - DriveLetter: $($Drive.DriveLetter), Path: $($Drive.Path), Name: $($Drive.Name)"
            }
        } else {
            Write-Log "No mapped network drives found."
        }

        # Collect installed printers
        Write-Log "Collecting installed printers."
        $Printers = Get-WmiObject -Class Win32_Printer | ForEach-Object {
            @{
                PrinterName = $_.Name
                DriverName = $_.DriverName
                PortName = $_.PortName
                Shared = $_.Shared
            }
        }
        if ($Printers) {
            $Inventory.Printers = $Printers
            foreach ($Printer in $Printers) {
                Write-Log "Printer - PrinterName: $($Printer.PrinterName), DriverName: $($Printer.DriverName), PortName: $($Printer.PortName), Shared: $($Printer.Shared)"
            }
        } else {
            Write-Log "No printers installed."
        }

        # Export to JSON
        Write-Log "Exporting inventory to $OutputFilePath."
        $Inventory | ConvertTo-Json -Depth 4 | Out-File -FilePath $OutputFilePath -Encoding UTF8

        # Verify file was created
        if (Test-Path $OutputFilePath) {
            Write-Log "Successfully exported inventory to $OutputFilePath."
        } else {
            Write-Log "Failed to export inventory to $OutputFilePath."
            exit 1
        }
    }
    elseif ($Mode -eq "Import") {
        # Find the most recent JSON file for the user and computer
        $ImportFile = Get-ChildItem -Path $NetworkPath -Filter "Inventory_$Computer_$User_*.json" | Sort-Object LastWriteTime -Descending | Select-Object -First 1
        if (-not $ImportFile) {
            Write-Log "No matching inventory JSON file found in $NetworkPath for $Computer and $User."
            exit 1
        }
        $ImportFilePath = $ImportFile.FullName
        Write-Log "Importing inventory from $ImportFilePath."

        # Read JSON file
        $Inventory = Get-Content -Path $ImportFilePath -Raw | ConvertFrom-Json

        # Import drive mappings
        Write-Log "Importing drive mappings."
        foreach ($Drive in $Inventory.DriveMappings) {
            $DriveLetter = $Drive.DriveLetter
            $Path = $Drive.Path
            try {
                # Check if drive is already mapped
                $ExistingDrive = Get-WmiObject -Class Win32_MappedLogicalDisk | Where-Object { $_.DeviceID -eq $DriveLetter -and $_.ProviderName -eq $Path }
                if ($ExistingDrive) {
                    Write-Log "Drive $DriveLetter already mapped to $Path. Skipping."
                } else {
                    New-PSDrive -Name ($DriveLetter -replace ":","") -PSProvider FileSystem -Root $Path -Scope Global -Persist -ErrorAction Stop
                    Write-Log "Mapped drive $DriveLetter to $Path."
                }
            }
            catch {
                Write-Log "Failed to map drive $DriveLetter to $Path : $_"
            }
        }

        # Import printers
        Write-Log "Importing printers."
        foreach ($Printer in $Inventory.Printers) {
            $PrinterName = $Printer.PrinterName
            $PortName = $Printer.PortName
            $DriverName = $Printer.DriverName
            try {
                # Check if printer already exists
                $ExistingPrinter = Get-WmiObject -Class Win32_Printer | Where-Object { $_.Name -eq $PrinterName }
                if ($ExistingPrinter) {
                    Write-Log "Printer $PrinterName already exists. Skipping."
                } else {
                    # Add printer (assumes driver is installed; otherwise, driver must be pre-installed)
                    Add-Printer -Name $PrinterName -DriverName $DriverName -PortName $PortName -ErrorAction Stop
                    Write-Log "Added printer $PrinterName with driver $DriverName on port $PortName."
                }
            }
            catch {
                Write-Log "Failed to add printer $PrinterName: $_"
            }
        }
    }

    Write-Log "$Mode process completed successfully."
    exit 0
}
catch {
    Write-Log "Error during $Mode process: $_"
    exit 1
}
finally {
    # Ensure log file is accessible
    if (Test-Path $LogFilePath) {
        Write-Log "Log file updated at $LogFilePath."
    } else {
        Write-Output "Failed to create or update log file at $LogFilePath."
        exit 1
    }
}