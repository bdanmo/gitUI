' vbscript

Function RunCmd(strCommand)
  ' run a command and show the output in a msgbox
  Dim objShell, objExecObject, strText
  Set objShell = CreateObject("WScript.Shell")
  Set objExecObject = objShell.Exec(strCommand)
  strText = objExecObject.StdOut.ReadAll
  Set objExecObject = Nothing
  Set objShell = Nothing
  RunCmd = strText
End Function

Sub Button1_Click()
  ' use RunCmd to run ipconfig, and show only the IPv4 address in a message box
  Dim strCommand, strOutput, strIP
  strCommand = "ipconfig"
  strOutput = RunCmd(strCommand)
  strIP = Mid(strOutput, InStr(strOutput, "IPv4 Address") + 36, 15)
  MsgBox strIP, vbInformation, "IPv4 Address"
End Sub

Sub Button2_Click()
  ' get hostname using cmd and show in a msgbox
  Dim strCommand, strOutput
  strCommand = "hostname"
  strOutput = RunCmd(strCommand)
  MsgBox strOutput, vbInformation, "Hostname"
End Sub

Sub Button3_Click()
  MsgBox "Button 3 clicked! Success!", vbInformation, "Success"
End Sub

Sub Button4_Click()
  MsgBox "Button 4 clicked! Success!", vbInformation, "Success"
End Sub