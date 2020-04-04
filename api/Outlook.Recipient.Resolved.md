---
title: Recipient.Resolved property (Outlook)
keywords: vbaol11.chm2352
f1_keywords:
- vbaol11.chm2352
ms.prod: outlook
api_name:
- Outlook.Recipient.Resolved
ms.assetid: 09c7655b-5acd-b527-56f6-59bc994a5ca1
ms.date: 06/08/2017
localization_priority: Normal
---


# Recipient.Resolved property (Outlook)

Returns a **Boolean** that indicates **True** if the recipient has been validated against the Address Book. Read-only.


## Syntax

_expression_. `Resolved`

_expression_ A variable that represents a '[Recipient](Outlook.Recipient.md)' object.


## Remarks

If similar names exist for a recipient in an Address Book, you can resolve the recipient by specifying the recipient's full SMTP email address.


## Example

This Visual Basic for Applications (VBA) example uses the  **[Resolve](Outlook.Recipient.Resolve.md)** method to resolve the **Recipient** object representing Dan Wilson, and then returns Dan's shared default **Calendar** folder.


```vb
Sub ResolveName() 
 
 Dim myNamespace As Outlook.NameSpace 
 
 Dim myRecipient As Outlook.Recipient 
 
 Dim CalendarFolder As Outlook.Folder 
 
 
 
 Set myNamespace = Application.GetNamespace("MAPI") 
 
 Set myRecipient = myNamespace.CreateRecipient("Dan Wilson") 
 
 myRecipient.Resolve 
 
 If myRecipient.Resolved Then 
 
 Call ShowCalendar(myNamespace, myRecipient) 
 
 End If 
 
End Sub 
 
 
 
Sub ShowCalendar(myNamespace, myRecipient) 
 
 Dim CalendarFolder As Outlook.Folder 
 
 Set CalendarFolder = _ 
 
 myNamespace.GetSharedDefaultFolder _ 
 
 (myRecipient, olFolderCalendar) 
 
 CalendarFolder.Display 
 
End Sub
```


## See also


[Recipient Object](Outlook.Recipient.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]