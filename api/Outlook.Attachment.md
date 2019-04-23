---
title: Attachment object (Outlook)
keywords: vbaol11.chm2360
f1_keywords:
- vbaol11.chm2360
ms.prod: outlook
api_name:
- Outlook.Attachment
ms.assetid: 3e11582b-ac90-0948-bc37-506570bb287b
ms.date: 06/08/2017
localization_priority: Normal
---


# Attachment object (Outlook)

Represents a document or link to a document contained in an Outlook item.


## Remarks

Use  **[Attachments](Outlook.Attachments.Item.md)** (_index_), where _index_ is the index number, to return a single **Attachment** object.

Use the  **[Add](Outlook.Attachments.Add.md)** method to add an attachment to an item.


## Example

The following Visual Basic for Applications (VBA) example creates a new mail message, attaches Q496.xlsx as an attachment (not a link), assigns the attachment a descriptive caption, and displays the mail message with this attachment. This example assumes that the specified spreadsheet, Q496.xlsx, exists in the specified path, D:\Documents.


```vb
Sub AddAttachment() 
 
 Dim myItem As Outlook.MailItem 
 
 Dim myAttachments As Outlook.Attachments 
 
 
 
 Set myItem = Application.CreateItem(olMailItem) 
 
 Set myAttachments = myItem.Attachments 
 
 myAttachments.Add "D:\Documents\Q496.xlsx", _ 
 
 olByValue, 1, "4th Quarter 1996 Results Chart" 
 
 myItem.Display 
 
End Sub
```


## See also


[Attach a File to a Mail Item](../outlook/How-to/Items-Folders-and-Stores/attach-a-file-to-a-mail-item.md)
[Attach an Outlook Contact Item to an Email Message](../outlook/Concepts/Attachments/attach-an-outlook-contact-item-to-an-email-message.md)
[Limit the Size of an Attachment to an Outlook Email Message](../outlook/Concepts/Attachments/limit-the-size-of-an-attachment-to-an-outlook-email-message.md)
[Modify an Attachment of an Outlook Email Message](../outlook/Concepts/Attachments/modify-an-attachment-of-an-outlook-email-message.md)
[Outlook Object Model Reference](overview/Outlook/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
