@echo off
@rem login office first
powershell -ExecutionPolicy Bypass -File "c:\\loginoffice.ps1"

@rem Run the entrypoint command specified via our command-line parameters
%*
