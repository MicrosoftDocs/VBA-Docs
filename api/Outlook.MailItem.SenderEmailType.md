---
title: MailItem.SenderEmailType property (Outlook)
keywords: vbaol11.chm1384
f1_keywords:
- vbaol11.chm1384
ms.prod: outlook
api_name:
- Outlook.MailItem.SenderEmailType
ms.assetid: e82cb8a6-d480-d1d1-ad15-a498ada6de37
ms.date: 06/08/2017
localization_priority: Normal
---


# MailItem.SenderEmailType property (Outlook)

Returns a  **String** that represents the type of entry for the email address of the sender of the Outlook item, such as 'SMTP' for Internet address, 'EX' for a Microsoft Exchange server address, etc. Read-only.


## Syntax

_expression_. `SenderEmailType`

_expression_ A variable that represents a [MailItem](Outlook.MailItem.md) object.


## Example

The following Microsoft Visual Basic for Applications (VBA) example demonstrates how to use the  **SenderEmailType** property. To run this example without errors, an email item should be open in the active inspector window.


```vb
Sub SenderEmailTypeExample() 
 
 Dim mail As Outlook.MailItem 
 
 
 
 Set mail = Application.ActiveInspector.CurrentItem 
 
 MsgBox mail.SenderEmailType 
 
 If mail.SenderEmailType = "SMTP" Then 
 
 MsgBox "Message from Internet email user." 
 
 Else 
 
 If mail.SenderEmailType = "EX" Then 
 
 MsgBox "Message from internal Exchange user." 
 
 End If 
 
 End If 
 
End Sub
```


## See also


[MailItem Object](Outlook.MailItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]