@echo off
echo import reg
@set baseDir="..\debug\bin"

regedit /s  E:\Work\docSword\plg\src\WordAddIn1\WordAddIn1\install\wpsReg.reg


C:\Windows\Microsoft.NET\Framework\v4.0.30319\RegAsm /codebase E:\Work\docSword\plg\src\WordAddIn1\WordAddIn1\bin\release\docSword.dll


@SET GACUTIL="E:\Work\docSword\plg\src\wpsDocSword\WpsWordAddin\install\NETFX 4.0 Tools\gacutil.exe"

Echo Install the dll into GAC
rem %GACUTIL% -i E:\Work\docSword\plg\src\wpsDocSword\WpsWordAddin\bin\Debug\WpsDocSword.dll
rem %GACUTIL% -i E:\Work\docSword\plg\src\wpsDocSword\WpsWordAddin\bin\Debug\word.dll
rem %GACUTIL% -i E:\Work\docSword\plg\src\wpsDocSword\WpsWordAddin\bin\Debug\office.dll
rem pause

pause



