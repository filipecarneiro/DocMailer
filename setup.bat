@echo off
REM Initial setup script for DocMailer on Windows

echo === DocMailer Initial Setup ===
echo.

REM Restore NuGet packages
echo Restoring NuGet packages...
dotnet restore

REM Build the project
echo Building the project...
dotnet build

REM Create directories if they don't exist
echo Creating required directories...
if not exist "Data" mkdir Data
if not exist "Output" mkdir Output
if not exist "Templates" mkdir Templates

echo.
echo === Setup Complete ===
echo.
echo Next steps:
echo 1. Edit config.json file with your email settings
echo 2. Place your Excel file in Data/ folder
echo 3. Run: dotnet run test (to test configuration)
echo 4. Run: dotnet run process (to process recipients)
echo.
pause
