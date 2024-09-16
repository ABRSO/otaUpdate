# Set ErrorActionPreference to stop the script on errors
$ErrorActionPreference = "Stop"

# Define path variables
$ZIP_FILE = "$env:TEMP\ota_files.zip"
$DEST_DIR = "C:\Project\aimlware\viaverde\exe"
$BACKUP_DIR = "${DEST_DIR}_backup"
$CONFIG_FILE = "C:\Project\aimlware\viaverde\config\config.json"
$UF2_FILE = "C:\Project\intern\otaUpdate\Blink_bootrom.ino.uf2"
$rp2040DriveLetter = "D:"
$TIMEOUT = 30

# Function Definitions

function ota_status_update {
    param (
        [string]$message,
        [string]$deviceid,
        [string]$ver
    )
try {
    $body = @{ message = $message }
    Invoke-RestMethod -Uri "https://brokerapi.viaverde.app/api/ota/update/$deviceid/version/$ver" `
        -Body (ConvertTo-Json $body) `
        -TimeoutSec $TIMEOUT `
        -Headers @{ "accept" = "application/json"; "Content-Type" = "application/json" } `
        -Method Patch
    }
    catch {
        Write-Host "StatusCode:" $_.Exception.Response.StatusCode.value__ 
        $streamReader = [System.IO.StreamReader]::new($_.Exception.Response.GetResponseStream())
        $ErrResp = $streamReader.ReadToEnd() | ConvertFrom-Json
        $streamReader.Close()
    }
    $ErrResp
}

# function ota_status_updateOG {
#     param (
#         [string]$message,
#         [string]$deviceid,
#         [string]$ver
#     )
#     $body = @{ message = $message } | ConvertTo-Json
#     Invoke-RestMethod -Uri "https://brokerapi.viaverde.app/api/ota/update/$deviceid/version/$ver" `
#         -Method Patch `
#         -TimeoutSec $TIMEOUT `
#         -Headers @{ "accept" = "application/json"; "Content-Type" = "application/json" } `
#         -Body $body
# }

function reload_and_restart_services {
    Write-Host "Reloading system configuration and restarting services"
    # Restart-Service -Name Spooler - check this one out
    Write-Host Restart-Service -NoNewWindow -FilePath "health_check" -Force -Name "health_check"
    Write-Host Restart-Service -NoNewWindow -FilePath "irrigation_controller" -Force -Name "irrigation_controller"
    #Start-Process -NoNewWindow -FilePath "sc.exe" -ArgumentList "stop irrigation_controller"
    #Start-Process -NoNewWindow -FilePath "sc.exe" -ArgumentList "start irrigation_controller"
    #Start-Process -NoNewWindow -FilePath "sc.exe" -ArgumentList "stop health_check"
    #Start-Process -NoNewWindow -FilePath "sc.exe" -ArgumentList "start health_check"
}

# Load the JSON config file and parse it
$config = Get-Content $CONFIG_FILE | ConvertFrom-Json
Write-Host "config: $config"

# Extract values from the JSON object
$deviceid = $config.deviceid
$device_type = $config.device_type
$batch_number = $config.batch_number
$software_version = $config.software_version

# Check if variables are empty
if (-not $deviceid) {
    Write-Host "Failed to extract deviceid from JSON file"
    exit 1
}
if (-not $device_type) {
    Write-Host "Failed to extract device_type from JSON file"
    exit 1
}
if (-not $batch_number) {
    Write-Host "Failed to extract batch_number from JSON file"
    exit 1
}
if (-not $software_version) {
    Write-Host "Failed to extract software_version from JSON file"
    exit 1
}

# Display the device parameters
Write-Host "Device parameters:"
Write-Host "deviceid: $deviceid"
Write-Host "device_type: $device_type"
Write-Host "batch_number: $batch_number"
Write-Host "software_version: $software_version"

# Form the API URL
$API_URL = "https://brokerapi.viaverde.app/api/ota/$software_version/type/$device_type/batch/$batch_number/deviceid/$deviceid"
Write-Host "API_URL: $API_URL"

# Send API Request and Process the Response
$API_RESPONSE = Invoke-RestMethod -Uri $API_URL -Method Get -TimeoutSec $TIMEOUT -Headers @{ "accept" = "application/json" }
# Write-Host "API_RESPONSE:" ($API_RESPONSE | ConvertTo-Json -Depth 10)

# Extract the status code from the response
$statusCode = $API_RESPONSE.statusCode
Write-Host "status code: $statusCode"
 
if ($statusCode -eq 200) {

    $data = $API_RESPONSE.data
    if (-not $data) {
        Write-Error "data is null or empty"
        ota_status_update "data is null or empty" $deviceid $software_version
        exit 1
    }

    $downloadUrl = $data.url
    if (-not $downloadUrl) {
        Write-Error "DOWNLOAD_URL is null or empty"
        ota_status_update "DOWNLOAD_URL is null or empty" $deviceid $software_version
        exit 1
    }
    if ($downloadUrl -eq "null") {
        Write-Error "No update found"
        ota_status_update "No update found" $deviceid $software_version
        exit 1
    }
}
else {
    Write-Error "Failed to obtain download URL"
    ota_status_update "Failed to obtain download URL" $deviceid $software_version
    exit 1
}

Write-Host "downloadURL: $downloadUrl"

Invoke-WebRequest -Uri $downloadUrl -OutFile $ZIP_FILE -UseBasicParsing -TimeoutSec $TIMEOUT
Write-Output "ZIP_FILE: $ZIP_FILE"

$TEMP_DIR = Join-Path $env:TEMP "ota_update_temp"
New-Item -Path $TEMP_DIR -ItemType Directory -Force
Write-Host "Temp: $TEMP_DIR"

try {
    Expand-Archive -Path $ZIP_FILE -DestinationPath $TEMP_DIR -Force
}
catch {
    Write-Error "Unzip failed"
    ota_status_update "Unzip failed" $deviceid $software_version

    # Remove the temporary directory
    Remove-Item -Path $TEMP_DIR -Recurse -Force
    exit 1
}

# Stop services before making the changes
Write-Output "Stopping services"
Write-Output "Stop-Service -Name "health_check" -Force"
Write-Output "Stop-Service -Name "irrigation_controller" -Force"

if (Test-Path $DEST_DIR) {
    # Backup the current directory
    Move-Item -Path $DEST_DIR -Destination $BACKUP_DIR -Force
    if (-not $?) {
        Write-Host "Failed to rename $DEST_DIR to $BACKUP_DIR"
        ota_status_update "Failed to rename $DEST_DIR to $BACKUP_DIR" $deviceid $software_version
        reload_and_restart_services
        exit 1
    }
    New-Item -Path $DEST_DIR -ItemType Directory -Force
}
else {
    Write-Output "$DEST_DIR does not exist, creating new directory"
    New-Item -Path $DEST_DIR -ItemType Directory -Force
}

# Copy new files to the destination directory
xcopy "$TEMP_DIR\*" "$DEST_DIR\" /E /I /Y
Get-ChildItem -Path $DEST_DIR -Recurse | Where-Object { $_.PSIsContainer -eq $false }  

if (-not $?) {
    Write-Output "Copy failed, rolling back"
    Remove-Item -Path $DEST_DIR -Recurse -Force

    if (Test-Path $BACKUP_DIR) {
        Move-Item -Path $BACKUP_DIR -Destination $DEST_DIR -Force
        if (-not $?) {
            Write-Output "Failed to rollback"
            ota_status_update "Copy failed and failed to rollback" $deviceid $software_version
            reload_and_restart_services
            exit 1
        }
    }
    else {
        Write-Output "Backup directory does not exist, cannot restore original state"
        ota_status_update "Copy failed. Backup directory does not exist, cannot rollback to previous version" $deviceid $software_version
        reload_and_restart_services
        exit 1
    }
    ota_status_update "Copy failed, rolled back to previous version" $deviceid $software_version
    reload_and_restart_services
    exit 1
}
else {
    # Remove backup directory if copy succeeds 
    Remove-Item -Path $BACKUP_DIR -Recurse -Force
}

# Set permissions of all files in DEST_DIR to -rwxrwxr-x
Write-Host "dest: $DEST_DIR"
icacls $DEST_DIR /grant "Everyone:(OI)(CI)F"
if (-not $?) {
    Write-Output "Failed to set permissions on $DEST_DIR"
    ota_status_update "Failed to set permissions on $DEST_DIR" $deviceid $software_version
    reload_and_restart_services
    exit 1
}

# Clean up
Remove-Item -Path $TEMP_DIR -Recurse -Force
if (Test-Path $ZIP_FILE) {
    Remove-Item -Force $ZIP_FILE
}

# Detect the COM port for RP2040
$deviceDescription = "USB Serial Device"

# Use Get-WMIObject to find the COM port by matching the device description
$comPort = Get-WmiObject Win32_SerialPort | Where-Object { $_.Description -like "*$deviceDescription*" } | Select-Object -ExpandProperty DeviceID
# Check if the COM port was found
if (-not $comPort) {
    Write-Host "RP2040 COM port not found. Exiting."
    exit 1
}
Write-Host "RP2040 detected on $comPort."

# Run the mode command to set baud rate directly from PowerShell
cmd /c "mode ${comPort}: baud=1200 parity=n data=8 stop=1"

Write-Host "Waiting for RP2040 drive to be detected..."

do {
    Start-Sleep -Seconds 1
} until (Test-Path "$rp2040DriveLetter\")

Write-Host "RP2040 drive detected at $rp2040DriveLetter."

# Copy the UF2 file to the RP2040 drive
try {
    Copy-Item -Path $UF2_FILE -Destination "$rp2040DriveLetter\"
    Write-Host "Copied UF2 file to $rp2040DriveLetter"
} catch {
    if (-not $UF2_FILE) {
        Write-Host "The provided UF2 file path is incorrect"
        exit 1
    }
    Write-Host "Failed to copy UF2 file. Error: $_"
    exit 1
}

# Updating config message and restarting services
Write-Output "Updating config"
Set-Location -Path $DEST_DIR
Write-Host Start-Process -FilePath "irrigation_controller.exe" -ArgumentList "version" -NoNewWindow -Wait
Set-Location "C:\Project\intern\otaUpdate\"

# Send Response with OTA status
$deviceid = $config.deviceid
$software_version = $config.software_version

Write-Output "deviceid: $deviceid"
Write-Output "software version: $software_version"
ota_status_update "Update successful" $deviceid $software_version

# Reload and restart irrigation controller and health check
reload_and_restart_services

Write-Output "Update completed successfully"
