---
title: Attachments object (Outlook)
keywords: vbaol11.chm169
f1_keywords:
- vbaol11.chm169
ms.prod: outlook
api_name:
- Outlook.Attachments
ms.assetid: 4cc96a5f-a822-8ad5-6f61-e996bee8ba22
ms.date: 06/08/2017
localization_priority: Normal
---


# Attachments object (Outlook)

Contains a set of  **[Attachment](Outlook.Attachment.md)** objects that represent the attachments in an Outlook item.


## Remarks

Use the  **[Attachments](Outlook.Attachments.Item.md)** property to return the **Attachments** collection for any Outlook item (except notes).

Use the  **[Add](Outlook.Attachments.Add.md)** method to add an attachment to an item.

To ensure consistent results, always save an item before adding or removing objects in the  **Attachments** collection of the item.


## Example

The following Visual Basic for Applications (VBA) example creates a new mail message, attaches a Q496.xls as an attachment (not a link), and gives the attachment a descriptive caption.


```vb
Set myItem = Application.CreateItem(olMailItem) 
 
myItem.Save 
 
Set myAttachments = myItem.Attachments 
 
myAttachments.Add "C:\My Documents\Q496.xls", _ 
 
 olByValue, 1, "4th Quarter 1996 Results Chart"
```


## Methods



|Name|
|:-----|
|[Add](Outlook.Attachments.Add.md)|
|[Item](Outlook.Attachments.Item.md)|
|[Remove](Outlook.Attachments.Remove.md)|

## Properties



|Name|
|:-----|
|[Application](Outlook.Attachments.Application.md)|
|[Class](Outlook.Attachments.Class.md)|
|[Count](Outlook.Attachments.Count.md)|
|[Parent](Outlook.Attachments.Parent.md)|
|[Session](Outlook.Attachments.Session.md)|

## See also


[Attach a File to a Mail Item](../outlook/How-to/Items-Folders-and-Stores/attach-a-file-to-a-mail-item.md)
[Attach an Outlook Contact Item to an Email Message](../outlook/Concepts/Attachments/attach-an-outlook-contact-item-to-an-email-message.md)
[Limit the Size of an Attachment to an Outlook Email Message](../outlook/Concepts/Attachments/limit-the-size-of-an-attachment-to-an-outlook-email-message.md)
[Modify an Attachment of an Outlook Email Message](../outlook/Concepts/Attachments/modify-an-attachment-of-an-outlook-email-message.md)
[Attachments Object Members](overview/Outlook.md)
[Outlook Object Model Reference](overview/Outlook/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
