---
title: DocumentItem.BeforeAttachmentRead event (Outlook)
ms.prod: outlook
api_name:
- Outlook.DocumentItem.BeforeAttachmentRead
ms.assetid: 22ed23a8-42a5-09bd-73b9-10591bfa7de9
ms.date: 06/08/2017
localization_priority: Normal
---


# DocumentItem.BeforeAttachmentRead event (Outlook)

Occurs before an attachment associated with an instance of the parent object is read from the file system, an attachment stream, or an  **[Attachment](Outlook.Attachment.md)** object.


## Syntax

_expression_. `BeforeAttachmentRead`( `_Attachment_` , `_Cancel_` )

_expression_ A variable that represents a [DocumentItem](Outlook.DocumentItem.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Attachment_|Required| **Attachment**|The  **Attachment** to be read.|
| _Cancel_|Required| **Boolean**|Set to  **True** to cancel the operation; otherwise, set to **False** to allow the **Attachment** to be read.|

## See also


[DocumentItem Object](Outlook.DocumentItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]