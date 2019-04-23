---
title: Prompt a user before closing a form
ms.prod: access
ms.assetid: 3a29f7c0-5692-49f0-bbfe-f9132d5b582f
ms.date: 09/25/2018
localization_priority: Normal
---


# Prompt a user before closing a form

The following example illustrates how to prompt the user to verify that the form should be closed.


```vb
Private Sub Form_Unload(Cancel As Integer) 
 If MsgBox("Are you sure that you want to close this form?", vbYesNo) = vbYes Then 
 Exit Sub 
 Else 
 Cancel = True 
 End If 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]