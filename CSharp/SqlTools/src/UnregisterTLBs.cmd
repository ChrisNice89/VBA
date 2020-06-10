@echo off

REM To use this batch file, run this file elevated (runas administrator)

SET TargetName=AccessCodeLib.Data.SqlTools.interop

SET BaseDir=%0
SET BaseDir=%BaseDir:\UnregisterTLBs.cmd=%

SET BinDir=%BaseDir%\SqlTools.interop\bin\Debug
SET TargetDir=%BinDir%

::SET NETPATH=%windir%\Microsoft.NET\Framework\v2.0.50727
SET NETPATH=%windir%\Microsoft.NET\Framework\v4.0.30319

ECHO Unregistering %TargetName%.dll
SET BinPath=%BinDir%\%TargetName%.dll
%NETPATH%\regasm.exe "%BinPath%" /codebase /tlb:"%TargetDir%\%TargetName%.tlb" /nologo /silent /unregister

pause