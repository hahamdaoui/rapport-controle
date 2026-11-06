@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"
title Rapport de controle - Serveur local

where python >nul 2>nul
if errorlevel 1 (
  echo [ERREUR] Python 3.11 ou superieur est introuvable.
  echo Installez-le depuis https://www.python.org/downloads/ ^(cochez "Add Python to PATH"^).
  pause
  exit /b 1
)

if not exist ".venv\Scripts\python.exe" (
  echo Premiere utilisation : creation de l'environnement et installation des modules...
  python -m venv .venv
  if errorlevel 1 ( echo [ERREUR] Creation de l'environnement impossible. & pause & exit /b 1 )
  ".venv\Scripts\python.exe" -m pip install --upgrade pip
  ".venv\Scripts\python.exe" -m pip install -r requirements.txt
  if errorlevel 1 ( echo [ERREUR] Installation des modules impossible ^(connexion Internet requise^). & pause & exit /b 1 )
)

if not exist "C:\Program Files\LibreOffice\program\soffice.exe" if not exist "C:\Program Files (x86)\LibreOffice\program\soffice.exe" (
  echo.
  echo [ATTENTION] LibreOffice n'a pas ete detecte : le PDF ne pourra pas etre genere.
  echo Installez-le gratuitement : https://www.libreoffice.org/download/
  echo.
)

echo Demarrage de l'application...
start "" cmd /c "timeout /t 3 /nobreak >nul & start http://127.0.0.1:5000"
".venv\Scripts\python.exe" app.py
pause
