@echo off
title AutoOCR
set "x=%~dp0"
cd /d "%x%"
powershell -NoExit -Command "& { Set-ExecutionPolicy RemoteSigned -Scope CurrentUser; .\Scripts\activate; & python .\readDocFiles.py;}"