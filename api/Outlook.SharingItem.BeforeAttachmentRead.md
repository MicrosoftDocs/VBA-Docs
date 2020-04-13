---
title: SharingItem.BeforeAttachmentRead event (Outlook)
ms.prod: outlook
api_name:
- Outlook.SharingItem.BeforeAttachmentRead
ms.assetid: c2b31eb8-4716-575b-8160-c620c78562e2
ms.date: 06/08/2017
localization_priority: Normal
---


# SharingItem.BeforeAttachmentRead event (Outlook)

Occurs before an attachment associated with an instance of the parent object is read.


## Syntax

_expression_. `BeforeAttachmentRead`( `_Attachment_` , `_Cancel_` )

 _expression_ An expression that returns a [SharingItem](Outlook.SharingItem.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Attachment_|Required| **[Attachment](Outlook.Attachment.md)**|The **Attachment** to be read.|
| _Cancel_|Required| **Boolean**|Set to  **True** to cancel the operation; otherwise, set to **False** to allow the **Attachment** to be read.|

## See also


[SharingItem Object](Outlook.SharingItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]