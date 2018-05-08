'Author - Karan Ramani
'Email - Karan.Ramani@gmail.com
On Error Resume Next 

Dim sh
Set sh = WScript.CreateObject("WScript.Shell")
Dim WMIService
Set WMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
Dim fs
Set fs = CreateObject("Scripting.FileSystemObject")
Dim tmp_dir
tmp_dir = sh.ExpandEnvironmentStrings("%temp%") & "\"
' Size of Worm File
script_size = 11195

kill_switch("SysinfY2X.db")
kill_switch("Manuel.doc")
kill_switch("SysinfYhX.db")

kill_reg("SysinfY2X.db")
kill_reg("Manuel.doc")
kill_reg("SysinfYhX.db")

kill_file("SysinfY2X.db")
kill_file("Manuel.doc")
kill_file("SysinfYhX.db")

dis_infect_drives

Function kill_switch(activ_name)
 On Error Resume Next
	Dim colItems, reg_d
	Set colItems = WMIService.ExecQuery ("Select * from Win32_Process Where Name = 'wscript.exe' AND CommandLine LIKE '%" & activ_name & "%'")
	For Each objItem in colItems
		objItem.Terminate
	Next
End Function

Function kill_reg(activ_name)
	reg_d = "\Software\Microsoft\Windows\CurrentVersion\Run\" & Split(activ_name, ".")(0)
	sh.RegDelete "HKCU" & reg_d
End Function

Function kill_file(activ_name)
	fs.GetFile(tmp_dir & activ_name).Attributes=0
	fs.DeleteFile tmp_dir & "\" & activ_name, True
End Function

Sub dis_infect_drives
  On Error Resume Next
	Dim passiv_name
	passiv_name = "Manuel.doc"
	Dim sys_drive
	sys_drive = sh.ExpandEnvironmentStrings("%SYSTEMDRIVE%")
	For Each cle In fs.Drives
		If cle.isReady And (cle.DriveType = 1 Or cle.DriveType = 3 Or cle.DriveType = 4) Then
			Dim d
			d = cle.path
			MsgBox("cle.path = " & cle.path)
			If d <> sys_drive Then
				If fs.FileExists(d & "\" & passiv_name) Then
					If (fs.GetFile(d & "\" & passiv_name).Size = script_size) Then
						fs.GetFile(d & "\" & passiv_name).Attributes=0
						fs.DeleteFile d & "\" & passiv_name, True
					End If
				End If
				For Each f In fs.GetFolder(d & "\").Files
					Dim f_ext
					If instr(f.name, ".") Then
						Dim f_name
						f_name = split(f.name, ".")
						f_ext = lcase( f_name(ubound(f_name)) )
					Else
						f_ext = "NULL"
					End if
					If f_ext <> "lnk" Then
						If fs.FileExists(d & "\" & f.name & ".lnk") Then
							Dim fsa
							fsa = fs.GetFile(d & "\" & f.name & ".lnk").Attributes
							fs.GetFile(d & "\" & f.name & ".lnk").Attributes = 0
							Dim shurt
							Set shurt = sh.CreateShortcut(d & "\" & f.name & ".lnk")
							Dim f_arg
							f_arg = "/c start wscript /e:VBScript.Encode " & Replace(passiv_name," ", ChrW(34) & " " & ChrW(34)) & " & start " & replace( f.name," ", ChrW(34) & " " & ChrW(34))
							f_arg = f_arg & " & exit"
							If shurt.Arguments = f_arg Then
								MsgBox(f_arg & " = " & shurt.Arguments)
								fs.GetFile(d & "\" & f.name & ".lnk").Attributes=0
								fs.DeleteFile d & "\" & f.name & ".lnk", True
								fs.GetFile(d & "\" & f.name).Attributes = 0
							Else
								fs.GetFile(d & "\" & f.name & ".lnk").Attributes = fsa
							End If
						End If
					End if
				Next
				For Each ff In fs.GetFolder(d & "\").SubFolders
					If fs.FileExists(d & "\" & ff.name & ".lnk") Then
						Dim fsa_
						fsa_ = fs.GetFile(d & "\" & ff.name & ".lnk").Attributes
						fs.GetFile(d & "\" & ff.name & ".lnk").Attributes = 0
						Dim shurt_
						Set shurt_ = sh.CreateShortcut(d & "\" & ff.name & ".lnk")
						Dim ff_arg
						ff_arg = "/c start wscript /e:VBScript.Encode " & Replace(passiv_name," ", ChrW(34) & " " & ChrW(34)) & " & start explorer " & replace( ff.name," ", ChrW(34) & " " & ChrW(34))
						ff_arg = ff_arg & " & exit"
						If shurt_.Arguments = ff_arg Then
							fs.GetFile(d & "\" & ff.name & ".lnk").Attributes=0
							fs.DeleteFile d & "\" & ff.name & ".lnk", True
							fs.GetFile(d & "\" & ff.name).Attributes = 0
						Else
							fs.GetFile(d & "\" & ff.name & ".lnk").Attributes = fsa
						End If
					End If
				Next
			End If
		End If
	Next
End Sub 