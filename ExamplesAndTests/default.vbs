
Sub Test()
  
   Dim oShell, scriptPath, appCmd, retVal, retData
   Set oShell  = CreateObject("WScript.Shell")
   scriptPath  = oShell.CurrentDirectory & "\Test.ps1"
   appCmd      = "powershell &'" & scriptPath & "'"
   retVal      = oShell.Run(appCmd, 4, true)
   retData     = window.clipboarddata.getdata("Text")
     
   If(retVal = 1) Then
      msgbox retData, vbExclamation, "Error"
   Else
      msgbox retData, vbInformation, "Return Value"
   End If
End Sub
  
Sub Initialize()
     
End Sub
