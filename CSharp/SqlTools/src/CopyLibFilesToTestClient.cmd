@echo off

SET SourceName=ACLibSqlTools
SET TargetName=AccessCodeLib.Data.SqlTools.interop

SET BaseDir=%0
SET BaseDir=%BaseDir:"=%
SET BaseDir=%BaseDir:\CopyLibFilesToTestClient.cmd=%

SET BinDir=%BaseDir%\SqlTools.interop\bin\Debug
SET TargetDir=%BinDir%


:: copy COM dll + tlb
copy /Y %TargetDir%\%TargetName%.dll %BaseDir%\..\tests\TestClient\lib\%TargetName%.dll
copy /Y %TargetDir%\%TargetName%.tlb %BaseDir%\..\tests\TestClient\lib\%TargetName%.tlb

:: copy net dlls
copy /Y %TargetDir%\AccessCodeLib.Data.SqlTools.dll %BaseDir%\..\tests\TestClient\lib\AccessCodeLib.Data.SqlTools.dll
copy /Y %TargetDir%\AccessCodeLib.Data.Common.Sql.dll %BaseDir%\..\tests\TestClient\lib\AccessCodeLib.Data.Common.Sql.dll
copy /Y %TargetDir%\AccessCodeLib.Data.SqlTools.Converter.dll %BaseDir%\..\tests\TestClient\lib\AccessCodeLib.Data.SqlTools.Converter.dll

:: 
pause