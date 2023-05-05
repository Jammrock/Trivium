@echo OFF

REM The file server FQDN
REM SET fqdn=Orrstown.com
SET fqdn=hs2.kehr.home

REM Name of the share
REM SET share=Shared
SET share=Music

REM Drive letter for the share
SET drvLtr=X:

REM Map the drive. The user's credentials will be used.
call net use %%drvLtr%% /delete
call net use %%drvLtr%% \\%%fqdn%%\%%share%% /PERSISTENT:YES

REM create the shortcut
powershell -nologo -noninteractive -noprofile -ExecutionPolicy bypass -file .\Install-Shortcuts.ps1 -Install
