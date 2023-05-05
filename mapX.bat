@echo OFF

REM The file server FQDN
SET fqdn=Domain.com

REM Name of the share
SET share=Shared

REM Drive letter for the share
SET drvLtr=X:

REM Map the drive. The user's credentials will be used.
call net use %%drvLtr%% /delete
call net use %%drvLtr%% \\%%fqdn%%\%%share%% /PERSISTENT:YES

REM create the shortcut
powershell -nologo -noninteractive -noprofile -ExecutionPolicy bypass -file .\Install-Shortcuts.ps1 -Install
