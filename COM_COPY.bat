@echo off

REM Step 1: Detect the COM port for RP2040
set "deviceDescription=USB Serial Device"

for /f "tokens=2 delims==" %%i in ('wmic path Win32_SerialPort where "Description like '%%%deviceDescription%%%'" get DeviceID /value 2^>nul') do (
    set "comPort=%%i"
)

REM Check if the COM port was found
if "%comPort%"=="" (
    echo RP2040 COM port not found. Exiting.
    pause
    exit /b 1
)

echo RP2040 detected on %comPort%.

Mode %comPort%: Baud=1200 Parity=N Data=8 Stop=1

REM Step 2: Configure the drive letter and UF2 file path
SET uf2FilePath="C:\Project\aimlware\viaverde\exe\Blink_bootrom.ino.uf2"
SET rp2040DriveLetter=D:

REM Step 3: Wait for the RP2040 drive to be detected

REM Wait for RP2040 to be detected
:waitloop
if exist %rp2040DriveLetter%\ (
    goto :detected
)
timeout /t 1 >nul
goto :waitloop

:detected
echo RP2040 drive detected

echo RP2040 drive detected at %rp2040DriveLetter%.

REM Step 4: Copy the UF2 file to the RP2040 drive
copy %uf2FilePath% %rp2040DriveLetter%
echo Copied UF2 file to %rp2040DriveLetter%