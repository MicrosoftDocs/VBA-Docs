---
title: MailItem.Saved property (Outlook)
keywords: vbaol11.chm1314
f1_keywords:
- vbaol11.chm1314
ms.prod: outlook
api_name:
- Outlook.MailItem.Saved
ms.assetid: 54a436a6-3da4-89d0-e1a6-db45c3732d95
ms.date: 06/08/2017
localization_priority: Normal
---


# MailItem.Saved property (Outlook)

Returns a **Boolean** value that is **True** if the Outlook item has not been modified since the last save. Read-only.


## Syntax

_expression_.**Saved**

_expression_ A variable that represents a [MailItem](Outlook.MailItem.md) object.


## Example

This Microsoft Visual Basic for Applications (VBA) example tests for the  **[Close](Outlook.MailItem.Close(even).md)** event and if the item has not been **[Saved](Outlook.MailItem.Saved.md)**, it uses the **[Save](Outlook.MailItem.Save.md)** method to save the item without prompting the user.


```vb
Public WithEvents myItem As Outlook.MailItem 
 
 
 
Public Sub Initialize_Handler() 
 
 Set myItem = Application.ActiveInspector.CurrentItem 
 
End Sub 
 
 
 
Private Sub myItem_Close(Cancel As Boolean) 
 
 If Not myItem.Saved Then 
 
 myItem.Save 
 
 MsgBox "Item was saved." 
 
 End If 
 
End Sub
```


## See also


[MailItem Object](Outlook.MailItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]