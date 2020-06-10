@echo off

SET BaseDir=%0
SET BaseDir=%BaseDir:"=%
SET BaseDir=%BaseDir:\CopyToTestClient.cmd=%

@echo %TestClientDir%

SET TestClientDir=%BaseDir%\..\..\..\..\tests\TestClient\lib
mkdir "%TestClientDir%"

copy "%BaseDir%\AccessCodeLib.Data.Common.Sql.dll" "%TestClientDir%\AccessCodeLib.Data.Common.Sql.dll"
copy "%BaseDir%\AccessCodeLib.Data.SqlTools.Converter.dll" "%TestClientDir%\AccessCodeLib.Data.SqlTools.Converter.dll"
copy "%BaseDir%\AccessCodeLib.Data.SqlTools.dll" "%TestClientDir%\AccessCodeLib.Data.SqlTools.dll"
copy "%BaseDir%\AccessCodeLib.Data.SqlTools.interop.dll" "%TestClientDir%\AccessCodeLib.Data.SqlTools.interop.dll"
copy "%BaseDir%\AccessCodeLib.Data.SqlTools.interop.tlb" "%TestClientDir%\AccessCodeLib.Data.SqlTools.interop.tlb"

pause
