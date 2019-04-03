---
title: DocumentItem.BeforeAttachmentPreview event (Outlook)
ms.prod: outlook
api_name:
- Outlook.DocumentItem.BeforeAttachmentPreview
ms.assetid: 687c0c41-c423-a30f-3fb6-562c2ab76f0c
ms.date: 06/08/2017
localization_priority: Normal
---


# DocumentItem.BeforeAttachmentPreview event (Outlook)

Occurs before an attachment associated with an instance of the parent object is previewed.


## Syntax

_expression_. `BeforeAttachmentPreview`( `_Attachment_` , `_Cancel_` )

_expression_ A variable that represents a [DocumentItem](Outlook.DocumentItem.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Attachment_|Required| **[Attachment](Outlook.Attachment.md)**|The  **Attachment** to be previewed.|
| _Cancel_|Required| **Boolean**|Set to  **True** to cancel the operation; otherwise, set to **False** to allow the **Attachment** to be previewed.|

## Remarks

This event occurs before an attachment is previewed, either from the attachment strip in the Reading Pane of the active explorer or from the active inspector.


## See also


[DocumentItem Object](Outlook.DocumentItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]