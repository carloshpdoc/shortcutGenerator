 Rem --- Criar Atalho MyAPP --- 

  Dim a,b,c
  ' Input Box with a Title
  a=InputBox("Deseja o MyAPP? Digite S ou N: ")
  'MsgBox ("Você escolheu: " & a)
  

	set Shell = WScript.CreateObject("WScript.Shell")
    LocalAtalho = Shell.SpecialFolders("C:\Users\%user%\Desktop")  
	strDesktop = WshShell.SpecialFolders("Desktop")

If StrComp(a, "s", 1) = 0 Then   

	LocalSoft = "\\SRV\WORK\MyAPP.exe"
	set lnk = Shell.CreateShortcut(strDesktop & "MyAPP.lnk")
	lnk.TargetPath = LocalSoft
	lnk.IconLocation = "\\SRV\WORK\MyAPP.exe, 0"
	lnk.WindowStyle = "1"
	lnk.Description = "MyAPP"	
	lnk.Save 
	MsgBox "Criado com sucesso o MyAPP em sua area de trabalho." 
 End If