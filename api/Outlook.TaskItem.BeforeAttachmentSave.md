---
title: TaskItem.BeforeAttachmentSave event (Outlook)
ms.prod: outlook
api_name:
- Outlook.TaskItem.BeforeAttachmentSave
ms.assetid: 93d31d5c-fb22-ce19-bcf2-651acc2d5db7
ms.date: 06/08/2017
localization_priority: Normal
---


# TaskItem.BeforeAttachmentSave event (Outlook)

Occurs just before an attachment is saved.


## Syntax

_expression_. `BeforeAttachmentSave`( `_Attachment_` , `_Cancel_` )

_expression_ A variable that represents a [TaskItem](Outlook.TaskItem.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Attachment_|Required| **[Attachment](Outlook.Attachment.md)**|The **Attachment** to be saved.|
| _Cancel_|Required| **Boolean**|(Not used in VBScript).  **False** when the event occurs. If the event procedure sets this argument to **True**, the save operation is not completed and the attachment is not changed.|

## Remarks

This event corresponds to when attachments are saved to the messaging store. The **BeforeAttachmentSave** event occurs just before an attachment is saved when an item is saved. If a user edits an attachment and then saves those changes, the **BeforeAttachmentSave** event will not occur at that time; instead it will occur when the item itself is later saved. It also does not occur when the attachment is saved on the hard disk using the **SaveAsFile** method.

In VBScript, if you set the return value of this function to  **False**, the save operation is cancelled and the attachment is not changed.


## See also


[TaskItem Object](Outlook.TaskItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]