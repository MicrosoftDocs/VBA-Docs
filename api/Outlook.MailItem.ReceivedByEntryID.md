---
title: MailItem.ReceivedByEntryID property (Outlook)
keywords: vbaol11.chm1341
f1_keywords:
- vbaol11.chm1341
ms.prod: outlook
api_name:
- Outlook.MailItem.ReceivedByEntryID
ms.assetid: db4325d3-4442-220d-a812-1d3e4a0085bf
ms.date: 06/08/2017
localization_priority: Normal
---


# MailItem.ReceivedByEntryID property (Outlook)

Returns a **String** representing the **[EntryID](Outlook.Recipient.EntryID.md)** for the true recipient as set by the transport provider delivering the mail message. Read-only.


## Syntax

_expression_. `ReceivedByEntryID`

_expression_ A variable that represents a [MailItem](Outlook.MailItem.md) object.


## Remarks

This property corresponds to the MAPI property  **PidTagReceivedByEntryId**.

If you are getting this property in a Microsoft Visual Basic or Microsoft Visual Basic for Applications (VBA) solution, owing to some type issues, instead of directly referencing  **ReceivedByEntryID**, you should get the property through the **[PropertyAccessor](Outlook.PropertyAccessor.md)** object returned by the **[MailItem.PropertyAccessor](Outlook.MailItem.PropertyAccessor.md)** property, specifying the **PidTagReceivedByEntryId** property and its MAPI proptag namespace. The following code sample in VBA shows the workaround.




```vb
Public Sub GetReceiverEntryID() 
 
 Dim objInbox As Outlook.Folder 
 
 Dim objMail As Outlook.MailItem 
 
 Dim oPA As Outlook.PropertyAccessor 
 
 Dim strEntryID As String 
 
 Const PidTagReceivedByEntryId As String = "http://schemas.microsoft.com/mapi/proptag/0x003F0102" 
 
 
 
 Set objInbox = Application.Session.GetDefaultFolder(olFolderInbox) 
 
 Set objMail = objInbox.Items(1) 
 
 Set oPA = objMail.PropertyAccessor 
 
 strEntryID = oPA.BinaryToString(oPA.GetProperty(PidTagReceivedByEntryId)) 
 
 Debug.Print strEntryID 
 
 
 
 Set objInbox = Nothing 
 
 Set objMail = Nothing 
 
End Sub
```


## See also


[MailItem Object](Outlook.MailItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]