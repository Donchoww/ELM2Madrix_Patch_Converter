@echo off
setlocal EnableExtensions

rem Dossier et script PowerShell
set "SCRIPT_DIR=%~dp0"
set "PS1=%SCRIPT_DIR%elm2madrix_full.ps1"

if not exist "%PS1%" (
  echo [ERROR] Introuvable: %PS1%
  pause
  exit /b 1
)

rem Choisir pwsh si dispo, sinon Windows PowerShell
where /Q pwsh.exe
if errorlevel 1 (
  set "PS=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"
) else (
  set "PS=pwsh.exe"
)

rem Console en UTF-8
chcp 65001 >NUL 2>&1

rem ---------------------- Parsing des arguments (robuste) ----------------------
rem - Garde uniquement les .csv ou dossiers existants
rem - Ignore le .ps1 si glissé par erreur
set "PATH_ARGS="
set "HAS_DIR=0"
set /a ARGCOUNT=0

:PARSE_ARGS
if "%~1"=="" goto AFTER_PARSE

if /I not "%~f1"=="%PS1%" (
  if /I "%~x1"==".csv" (
    if exist "%~f1" (
      call :APPEND_PATH "%~f1"
      set /a ARGCOUNT+=1
    )
  ) else (
    if exist "%~f1\" (
      call :APPEND_PATH "%~f1"
      set "HAS_DIR=1"
      set /a ARGCOUNT+=1
    )
  )
)
shift
goto PARSE_ARGS

:AFTER_PARSE
if "%HAS_DIR%"=="1" (
  set "RECURSE_SWITCH=-Recurse"
) else (
  set "RECURSE_SWITCH="
)

echo [INFO ] Lancement de %PS1%
echo [INFO ] %ARGCOUNT% argument(s) retenu(s) après filtre.
echo [INFO ] Args = %RECURSE_SWITCH% -Paths %PATH_ARGS%
rem ------------------------------------------------------------------------------

"%PS%" -NoProfile -ExecutionPolicy Bypass -NoLogo -File "%PS1%" %RECURSE_SWITCH% -Paths %PATH_ARGS%
set "RC=%ERRORLEVEL%"

echo(
if "%RC%"=="0" (
  echo [ OK  ] Terminé avec succes.
) else (
  echo [ERROR] Terminé avec erreurs. Code=%RC%
)
pause
exit /b %RC%

:APPEND_PATH
rem Conserver les guillemets -> utiliser %1 (pas %~1)
set "PATH_ARGS=%PATH_ARGS% %1"
goto :eof
