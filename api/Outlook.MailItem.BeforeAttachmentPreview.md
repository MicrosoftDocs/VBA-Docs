---
title: MailItem.BeforeAttachmentPreview event (Outlook)
ms.prod: outlook
api_name:
- Outlook.MailItem.BeforeAttachmentPreview
ms.assetid: 279e1af4-38e1-d6b5-50a5-9ebd517826ae
ms.date: 06/08/2017
localization_priority: Normal
---


# MailItem.BeforeAttachmentPreview event (Outlook)

Occurs before an attachment associated with an instance of the parent object is previewed.


## Syntax

_expression_. `BeforeAttachmentPreview`( `_Attachment_` , `_Cancel_` )

_expression_ A variable that represents a [MailItem](Outlook.MailItem.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Attachment_|Required| **[Attachment](Outlook.Attachment.md)**|The **Attachment** to be previewed.|
| _Cancel_|Required| **Boolean**|Set to  **True** to cancel the operation; otherwise, set to **False** to allow the **Attachment** to be previewed.|

## Remarks

This event occurs before an attachment is previewed, either from the attachment strip in the Reading Pane of the active explorer or from the active inspector.


## See also


[MailItem Object](Outlook.MailItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]