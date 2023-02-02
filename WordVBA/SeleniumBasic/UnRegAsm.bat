@ECHO OFF
cd /d %~dp0
"C:\Windows\Microsoft.NET\Framework\v4.0.30319\RegAsm.exe" /u "SeleniumBasic.dll" /tlb "SeleniumBasic.tlb"
"C:\Windows\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe" /u "SeleniumBasic.dll" /tlb "SeleniumBasic.tlb"
PAUSE
CLS