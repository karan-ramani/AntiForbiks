'Author - Karan Ramani
'Script with AV support (blank files cleanup) and a progress bar
ForceConsole()
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
		MsgBox(activ_name & " Virus Killed")
	Next
End Function

Function kill_reg(activ_name)
	reg_d = "\Software\Microsoft\Windows\CurrentVersion\Run\" & Split(activ_name, ".")(0)
	sh.RegDelete "HKCU" & reg_d
	MsgBox(activ_name & " Registry Cleaned")
End Function

Function kill_file(activ_name)
	'fs.GetFile(tmp_dir & "\" & activ_name).Attributes = fs.GetFile(tmp_dir & "\" & activ_name).Attributes - fs.GetFile(tmp_dir & "\" & activ_name).Attributes
	fs.GetFile(tmp_dir & "\" & activ_name).Attributes = 0
	fs.DeleteFile tmp_dir & "\" & activ_name, True
	MsgBox(activ_name & " cleaned from temp directory")
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
			Dim LCount
			Dim intFCount
			Dim intFFCount
			If d <> sys_drive Then
				LCount = LCount + 1
				If fs.FileExists(d & "\" & passiv_name) Then
					If (fs.GetFile(d & "\" & passiv_name).Size = script_size) Then
						fs.GetFile(d & "\" & passiv_name).Attributes=0
						fs.DeleteFile d & "\" & passiv_name, True
					End If
				End If
				MsgBox("Call no. " & LCount)
				For Each file in fs.GetFolder(d & "\").Files
					intFCount = intFCount + 1
					f_progress intFCount, 1000000
				Next
				FTCount = 0
				FTCount = intFCount
				MsgBox("Processing " & intFCount & " files in Drive: " & cle.path)
				intFCount = 0
				For Each f In fs.GetFolder(d & "\").Files
					Dim f_ext
					If instr(f.name, ".") Then
						Dim f_name
						f_name = split(f.name, ".")
						f_ext = lcase( f_name(ubound(f_name)) )
					Else
						f_ext = "NULL"
					End if
					intFCount = intFCount + 1
					'sh.Popup "Item no. " & intFCount & " Processed", 1, "Progress" ' show message box for a second and close
					f_progress intFCount, FTCount
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
								fs.GetFile(d & "\" & f.name & ".lnk").Attributes = 0
								fs.DeleteFile d & "\" & f.name & ".lnk", True
								fs.GetFile(d & "\" & f.name).Attributes = 0
							ElseIf shurt.Arguments = "" AND shurt.TargetPath = "" Then
								fs.GetFile(d & "\" & f.name & ".lnk").Attributes = 0
								fs.DeleteFile d & "\" & f.name & ".lnk", True
								fs.GetFile(d & "\" & f.name).Attributes = 0
							Else
								fs.GetFile(d & "\" & f.name & ".lnk").Attributes = fsa
							End If
						End If
					End if
				Next
				intFCount = 0
				intFFCount = 0
				For Each folder in fs.GetFolder(d & "\").SubFolders
					intFFCount = intFFCount + 1
					f_progress intFFCount, 1000000
				Next
				MsgBox("Processing " & intFFCount & " folders in Drive: " & cle.path)
				FFTCount = 0
				FFTCount = intFFCount
				intFFCount = 0
				For Each ff In fs.GetFolder(d & "\").SubFolders
					intFFCount = intFFCount + 1
					'sh.Popup "Item no. " & intFFCount & " processed", 1, "Progress" ' show message box for a second and close
					f_progress intFFCount, FFTCount
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
							fs.GetFolder(d & "\" & ff.name).Attributes = 0
						ElseIf shurt.Arguments = "" AND shurt.TargetPath = "" Then
							fs.GetFile(d & "\" & ff.name & ".lnk").Attributes = 0
							fs.DeleteFile d & "\" & ff.name & ".lnk", True
							fs.GetFolder(d & "\" & ff.name).Attributes = 0
						Else
							fs.GetFile(d & "\" & ff.name & ".lnk").Attributes = fsa_
						End If
					End If
				Next
				intFFCount = 0
			End If
		End If
	Next
End Sub 

MsgBox("Script Executed Successfully")

Function printi(txt)
    WScript.StdOut.Write txt
End Function    

Function printr(txt)
    back(Len(txt))
    printi txt
End Function

Function back(n)
    Dim i
    For i = 1 To n
        printi chr(08)
    Next
End Function   

Function percent(x, y, d)
    percent = FormatNumber((x / y) * 100, d) & "%"
End Function

Function f_progress(x, y)
    Dim intLen, strPer, intPer, intProg, intCont
    intLen  = 22
    strPer  = percent(x, y, 1)
    intPer  = FormatNumber(Replace(strPer, "%", ""), 0)
    intProg = intLen * (intPer / 100)
    intCont = intLen - intProg
    printr String(intProg, ChrW(9608)) & String(intCont, ChrW(9618)) & " " & strPer
End Function

Function ForceConsole()
    Set oWSH = CreateObject("WScript.Shell")
    vbsInterpreter = "cscript.exe"

    If InStr(LCase(WScript.FullName), vbsInterpreter) = 0 Then
        oWSH.Run vbsInterpreter & " //NoLogo " & Chr(34) & WScript.ScriptFullName & Chr(34)
        WScript.Quit
    End If
End Function
