@echo off
setlocal

cd /d "%~dp0"
set "MSBUILD=C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\amd64\MSBuild.exe"
set "NAVISWORKS_VERSION=%~1"
if "%NAVISWORKS_VERSION%"=="" set "NAVISWORKS_VERSION=2026"
set "NAVISWORKS_DIR=C:\Program Files\Autodesk\Navisworks Manage %NAVISWORKS_VERSION%"
set "NAVISWORKS_PLUGIN_DIR=%NAVISWORKS_DIR%\Plugins\PROP"

if not exist "%MSBUILD%" (
  echo [ERRO] MSBuild nao encontrado em:
  echo   "%MSBUILD%"
  exit /b 1
)

if not exist "%NAVISWORKS_DIR%" (
  echo [ERRO] Navisworks %NAVISWORKS_VERSION% nao encontrado em:
  echo   "%NAVISWORKS_DIR%"
  exit /b 1
)

"%MSBUILD%" "PROP.csproj" /p:NavisworksVersion=%NAVISWORKS_VERSION% /p:Configuration=Release /p:Platform=AnyCPU /p:UseSharedCompilation=false /v:minimal
if errorlevel 1 (
  echo [ERRO] Falha na compilacao para Navisworks %NAVISWORKS_VERSION%
  exit /b %errorlevel%
)

copy /Y "NexusRibbon.xaml" "%NAVISWORKS_PLUGIN_DIR%\NexusRibbon.xaml" >nul
if errorlevel 1 (
  echo [ERRO] Falha ao atualizar NexusRibbon.xaml em:
  echo   "%NAVISWORKS_PLUGIN_DIR%"
  exit /b %errorlevel%
)

echo [OK] Plugin PROP instalado em: %NAVISWORKS_PLUGIN_DIR%
endlocal
