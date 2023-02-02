@ECHO OFF
cd /d %~dp0
"C:\Windows\Microsoft.NET\Framework\v4.0.30319\RegAsm.exe" /codebase "Selenium.dll" /tlb "Selenium32.tlb"
"C:\Windows\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe" /codebase "Selenium.dll" /tlb "Selenium64.tlb"
PAUSE
CLS