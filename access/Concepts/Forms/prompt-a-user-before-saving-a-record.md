---
title: Prompt a user before saving a record
ms.prod: access
ms.assetid: 4b47967c-a043-cc8a-774f-1df0b529f29b
ms.date: 09/25/2018
localization_priority: Normal
---


# Prompt a user before saving a record

The following example illustrates how to use the [BeforeUpdate](../../../api/Access.Form.BeforeUpdate(even).md) event to prompt users to confirm their changes each time they save a record in a form.


```vb
Private Sub Form_BeforeUpdate(Cancel As Integer) 
   Dim strMsg As String 
   Dim iResponse As Integer 
 
   ' Specify the message to display. 
   strMsg = "Do you wish to save the changes?" & Chr(10) 
   strMsg = strMsg & "Click Yes to Save or No to Discard changes." 
 
   ' Display the message box. 
   iResponse = MsgBox(strMsg, vbQuestion + vbYesNo, "Save Record?") 
    
   ' Check the user's response. 
   If iResponse = vbNo Then 
    
      ' Undo the change. 
      DoCmd.RunCommand acCmdUndo 
 
      ' Cancel the update. 
      Cancel = True 
   End If 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
