---
title: DistListItem.BeforeAttachmentRead event (Outlook)
ms.prod: outlook
api_name:
- Outlook.DistListItem.BeforeAttachmentRead
ms.assetid: f7c6f477-9f50-f099-eec4-67d12d4ca398
ms.date: 06/08/2017
localization_priority: Normal
---


# DistListItem.BeforeAttachmentRead event (Outlook)

Occurs before an attachment associated with an instance of the parent object is read from the file system, an attachment stream, or an  **[Attachment](Outlook.Attachment.md)** object.


## Syntax

_expression_. `BeforeAttachmentRead`( `_Attachment_` , `_Cancel_` )

_expression_ A variable that represents a [DistListItem](Outlook.DistListItem.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Attachment_|Required| **Attachment**|The  **Attachment** to be read.|
| _Cancel_|Required| **Boolean**|Set to  **True** to cancel the operation; otherwise, set to **False** to allow the **Attachment** to be read.|

## See also


[DistListItem Object](Outlook.DistListItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]