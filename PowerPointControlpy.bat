@echo off
pushd %~dp0
start "%~n0" python %~n0.py
popd
