---
title: DistListItem.BeforeAttachmentAdd event (Outlook)
ms.prod: outlook
api_name:
- Outlook.DistListItem.BeforeAttachmentAdd
ms.assetid: 96e8cd2d-fabb-f026-c442-483f46414a21
ms.date: 06/08/2017
localization_priority: Normal
---


# DistListItem.BeforeAttachmentAdd event (Outlook)

Occurs before an attachment is added to an instance of the parent object.


## Syntax

_expression_. `BeforeAttachmentAdd`( `_Attachment_` , `_Cancel_` )

_expression_ A variable that represents a [DistListItem](Outlook.DistListItem.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Attachment_|Required| **[Attachment](Outlook.Attachment.md)**|The **Attachment** to be added to the item.|
| _Cancel_|Required| **Boolean**|Set to  **True** to cancel the operation; otherwise, set to **False** to allow the **Attachment** to be added.|

## See also


[DistListItem Object](Outlook.DistListItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]