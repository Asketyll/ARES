@echo off
:: ============================================================
::  ARES Multi-Agent Launcher
::  IMPORTANT : lancer en double-clic NORMAL (pas admin)
:: ============================================================

set PROJECT=C:\Dev\ARES

:: Verifie que le projet existe
if not exist "%PROJECT%" (
    echo ERREUR: Dossier %PROJECT% introuvable
    pause
    exit /b 1
)

echo Lancement ARES Multi-Agent...

wt.exe new-tab --title "Architecte" --tabColor "#1a472a" cmd /k "cd /d %PROJECT% && echo === AGENT ARCHITECTE === && echo. && claude" ; new-tab --title "Developpeur" --tabColor "#1e3a5f" cmd /k "cd /d %PROJECT% && echo === AGENT DEVELOPPEUR === && echo. && claude" ; new-tab --title "Reviewer" --tabColor "#4a1942" cmd /k "cd /d %PROJECT% && echo === AGENT REVIEWER === && echo. && claude"

echo Fini !
timeout /t 2 >nul