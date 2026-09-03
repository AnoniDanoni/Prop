@echo off
setlocal EnableExtensions EnableDelayedExpansion

cd /d "%~dp0"
set "MSBUILD=C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\amd64\MSBuild.exe"
set "PATCH_ROOT=%~dp0..\build\PROP"
set BUILD_COUNT=0

if not exist "%MSBUILD%" (
  echo [ERRO] MSBuild nao encontrado em:
  echo   "%MSBUILD%"
  exit /b 1
)

call :CREATE_PATCH_DIR
if errorlevel 1 exit /b %errorlevel%

call :BUILD_VERSION 2026
if errorlevel 1 exit /b %errorlevel%

call :BUILD_VERSION 2024
if errorlevel 1 exit /b %errorlevel%

if %BUILD_COUNT% EQU 0 (
  echo [ERRO] Nenhuma versao do Navisworks 2024/2026 foi encontrada.
  exit /b 1
)

echo Build completo. %BUILD_COUNT% versao(oes) compilada(s).
echo Patch gerado em: %PATCH_DIR%
endlocal
exit /b 0

:BUILD_VERSION
set "NAVISWORKS_VERSION=%~1"
set "NAVISWORKS_DIR=C:\Program Files\Autodesk\Navisworks Manage %NAVISWORKS_VERSION%"

if not exist "%NAVISWORKS_DIR%" (
  echo [AVISO] Navisworks %NAVISWORKS_VERSION% nao encontrado - pulando.
  exit /b 0
)

echo Compilando PROP para Navisworks %NAVISWORKS_VERSION%...
"%MSBUILD%" "PROP.csproj" /p:NavisworksVersion=%NAVISWORKS_VERSION% /p:PatchOutputDir="%PATCH_DIR%\%NAVISWORKS_VERSION%" /p:Configuration=Release /p:Platform=AnyCPU /p:UseSharedCompilation=false /v:minimal
if errorlevel 1 (
  echo [ERRO] Falha na compilacao para Navisworks %NAVISWORKS_VERSION%
  exit /b %errorlevel%
)

set /a BUILD_COUNT+=1
echo   [OK] Plugin instalado em: %NAVISWORKS_DIR%\Plugins\PROP\
echo.
exit /b 0

:CREATE_PATCH_DIR
set "LAST_PATCH=-1"

if not exist "%PATCH_ROOT%" mkdir "%PATCH_ROOT%"
if errorlevel 1 (
  echo [ERRO] Nao foi possivel criar a pasta de patch:
  echo   "%PATCH_ROOT%"
  exit /b 1
)

for /f "delims=" %%D in ('dir /ad /b "%PATCH_ROOT%\1.0.*" 2^>nul') do (
  for /f "tokens=1,2,3 delims=." %%A in ("%%D") do (
    if "%%A.%%B"=="1.0" (
      set "PATCH_NUMBER=%%C"
      for /f "delims=0123456789" %%N in ("!PATCH_NUMBER!") do set "PATCH_NUMBER="
      if defined PATCH_NUMBER (
        set /a CURRENT_PATCH=!PATCH_NUMBER!
        if !CURRENT_PATCH! GTR !LAST_PATCH! set "LAST_PATCH=!CURRENT_PATCH!"
      )
    )
  )
)

set /a NEXT_PATCH=LAST_PATCH+1
set "PATCH_DIR=%PATCH_ROOT%\1.0.!NEXT_PATCH!"
mkdir "!PATCH_DIR!"
if errorlevel 1 (
  echo [ERRO] Nao foi possivel criar a versao de patch:
  echo   "!PATCH_DIR!"
  exit /b 1
)

exit /b 0
