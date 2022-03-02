:: Copyright (c) Microsoft Corporation.
:: Licensed under the MIT License.

@if "%SCM_TRACE_LEVEL%" NEQ "4" @echo off

:: ----------------------
:: KUDU Deployment Script
:: Version: 1.0.17
:: ----------------------

:: Prerequisites
:: -------------

:: Verify node.js installed
where node 2>nul >nul
IF %ERRORLEVEL% NEQ 0 (
  echo Missing node.js executable, please install node.js, if already installed make sure it can be reached from current environment.
  goto error
)

:: Setup
:: -----

setlocal enabledelayedexpansion

SET ARTIFACTS=%~dp0%..\artifacts

IF NOT DEFINED DEPLOYMENT_SOURCE (
  SET DEPLOYMENT_SOURCE=%~dp0%.
)

IF NOT DEFINED DEPLOYMENT_TARGET (
  SET DEPLOYMENT_TARGET=%ARTIFACTS%\wwwroot
)

IF NOT DEFINED NEXT_MANIFEST_PATH (
  SET NEXT_MANIFEST_PATH=%ARTIFACTS%\manifest

  IF NOT DEFINED PREVIOUS_MANIFEST_PATH (
    SET PREVIOUS_MANIFEST_PATH=%ARTIFACTS%\manifest
  )
)

IF NOT DEFINED KUDU_SYNC_CMD (
  :: Install kudu sync
  echo Installing Kudu Sync
  call npm install kudusync -g --silent
  IF !ERRORLEVEL! NEQ 0 goto error

  :: Locally just running "kuduSync" would also work
  SET KUDU_SYNC_CMD=%appdata%\npm\kuduSync.cmd
)
IF NOT DEFINED DEPLOYMENT_TEMP (
  SET DEPLOYMENT_TEMP=%temp%\___deployTemp%random%
  SET CLEAN_LOCAL_DEPLOYMENT_TEMP=true
)

IF DEFINED CLEAN_LOCAL_DEPLOYMENT_TEMP (
  IF EXIST "%DEPLOYMENT_TEMP%" rd /s /q "%DEPLOYMENT_TEMP%"
  mkdir "%DEPLOYMENT_TEMP%"
)

IF DEFINED MSBUILD_PATH goto MsbuildPathDefined
SET MSBUILD_PATH=%ProgramFiles(x86)%\MSBuild\14.0\Bin\MSBuild.exe
:MsbuildPathDefined
::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
:: Deployment
:: ----------

@REM echo Handling ASP.NET Core Web Application deployment.

:: 1. Restore nuget packages
@REM echo Restoring NuGet packages
@REM call :ExecuteCmd dotnet restore "%DEPLOYMENT_SOURCE%\Source\Microsoft.Teams.Apps.CompanyCommunicator.sln"
@REM IF !ERRORLEVEL! NEQ 0 goto error

:: 2. Restore npm packages
echo Restoring npm packages 1 (this can take several minutes)
echo "%DEPLOYMENT_SOURCE%"
pushd "%DEPLOYMENT_SOURCE%"

echo Restoring npm packages 2 (this can take several minutes)
echo "%DEPLOYMENT_SOURCE%\tabs"
pushd "%DEPLOYMENT_SOURCE%\tabs"

call :ExecuteCmd npm install --no-audit
IF !ERRORLEVEL! NEQ 0 (
    echo First attempt failed, retrying once
    call :ExecuteCmd npm install --no-audit
)
popd
IF !ERRORLEVEL! NEQ 0 goto error

:: 3. Build the client app
echo Building the client app (this can take several minutes)
pushd "%DEPLOYMENT_SOURCE%\tabs"
@REM call :ExecuteCmd npm run build
call :ExecuteCmd npm run-script build
popd
IF !ERRORLEVEL! NEQ 0 goto error

:: 4. Build and publish
echo Building the application
call :ExecuteCmd dotnet publish "%DEPLOYMENT_SOURCE%\tabs\build" --output "%DEPLOYMENT_TEMP%" --configuration Release -property:KuduDeployment=1
IF !ERRORLEVEL! NEQ 0 goto error

:: 5. KuduSync
call :ExecuteCmd "%KUDU_SYNC_CMD%" -v 50 -f "%DEPLOYMENT_SOURCE%\tabs\build" -t "%DEPLOYMENT_TARGET%" -n "%NEXT_MANIFEST_PATH%" -p "%PREVIOUS_MANIFEST_PATH%" -i ".git;.hg;.deployment;deploy.cmd"
IF !ERRORLEVEL! NEQ 0 goto error

::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
goto end

:: Execute command routine that will echo out when error
:ExecuteCmd
setlocal
set _CMD_=%*
call %_CMD_%
if "%ERRORLEVEL%" NEQ "0" echo Failed exitCode=%ERRORLEVEL%, command=%_CMD_%
exit /b %ERRORLEVEL%

:error
endlocal
echo An error has occurred during web site deployment.
call :exitSetErrorLevel
call :exitFromFunction 2>nul

:exitSetErrorLevel
exit /b 1

:exitFromFunction
()

:end
endlocal
echo Finished successfully.

