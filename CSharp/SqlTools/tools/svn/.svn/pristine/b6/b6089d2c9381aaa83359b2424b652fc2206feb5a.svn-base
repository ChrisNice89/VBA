@echo off

SET BaseDir=%0
SET BaseDir=%BaseDir:"=%
SET BaseDir=%BaseDir:\SubWCRev.cmd=%


if "%PROCESSOR_ARCHITECTURE%"=="x86" (  
   %BaseDir%\SubWCRev_x86.exe %1 %2 %3
) else ( 
   %BaseDir%\SubWCRev.exe %1 %2 %3
)
