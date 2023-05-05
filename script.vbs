' vbscript

Sub Button1_Click()
  ' write command to ipconfig and show current ONLY the ip address in a new msgbox
  Dim objShell, objExecObject, strText
  Set objShell = CreateObject("WScript.Shell")
  Set objExecObject = objShell.Exec("ipconfig")
  strText = objExecObject.StdOut.ReadAll
  Set objExecObject = Nothing
  Set objShell = Nothing
  Dim strIP
  strIP = Mid(strText, InStr(strText, "IPv4 Address") + 36, 15)
  MsgBox strIP, vbInformation, "IP Address"
  
End Sub

Sub Button2_Click()
  MsgBox "Button 2 clicked! Success!", vbInformation, "Success"
End Sub

Sub Button3_Click()
  MsgBox "Button 3 clicked! Success!", vbInformation, "Success"
End Sub

Sub Button4_Click()
  MsgBox "Button 4 clicked! Success!", vbInformation, "Success"
End Sub