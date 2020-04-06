---
title: NameSpace.GetRecipientFromID method (Outlook)
keywords: vbaol11.chm764
f1_keywords:
- vbaol11.chm764
ms.prod: outlook
api_name:
- Outlook.NameSpace.GetRecipientFromID
ms.assetid: 8475e869-ce1f-cd10-0c02-79a6dd5f9a8e
ms.date: 06/08/2017
localization_priority: Normal
---


# NameSpace.GetRecipientFromID method (Outlook)

Returns the  **[Recipient](Outlook.Recipient.md)** object that is identified by the specified entry ID (if valid).


## Syntax

_expression_. `GetRecipientFromID`( `_EntryID_` )

_expression_ A variable that represents a '[NameSpace](Outlook.NameSpace.md)' object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _EntryID_|Required| **String**|The  **[EntryID](Outlook.Recipient.EntryID.md)** of the recipient.|

## Return value

A  **Recipient** object that represents the specified recipient.


## Remarks

This method is used to ease the transition between MAPI and OLE/Messaging applications and Microsoft Outlook.

This method is similar to the  **[GetRecipientFromID](Outlook.Account.GetRecipientFromID.md)** method of the **[Account](Outlook.Account.md)** object. If there are multiple Microsoft Exchange accounts in the current profile, use the **GetRecipientFromID** method for the corresponding account.


## Example

This Visual Basic for Applications (VBA) example gets the entry ID of the first recipient for the first item in the  **[Items](Outlook.Items.md)** collection for the **Inbox** folder, uses **GetRecipientFromID** to obtain the recipient from the entry ID, and displays the recipient name.


```vb
Public Sub GetRecipientFromID() 
 
 Dim inbox As Outlook.folder 
 
 Dim mail As Outlook.MailItem 
 
 Dim rcp As Outlook.Recipient 
 
 Dim rcp1 As Outlook.Recipient 
 
 Dim strEntryId As String 
 
 
 
 'Get Inbox folder. 
 
 Set inbox = Application.session.GetDefaultFolder(olFolderInbox) 
 
 
 
 ' Get the first item in the Inbox. 
 
 Set mail = inbox.items(1) 
 
 
 
 ' Get the first recipient on that first item. 
 
 Set rcp = mail.Recipients.Item(1) 
 
 
 
 ' Get the string entry ID of the recipient. 
 
 strEntryId = rcp.entryID 
 
 
 
 ' Get the Recipient object based on the string entry ID. 
 
 Set rcp1 = Application.session.GetRecipientFromID(strEntryId) 
 
 
 
 MsgBox "Recipient Name: " & rcp1.name, vbInformation 
 
End Sub
```


## See also


[NameSpace Object](Outlook.NameSpace.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]