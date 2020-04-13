---
title: TaskRequestUpdateItem.BeforeAttachmentAdd event (Outlook)
ms.prod: outlook
api_name:
- Outlook.TaskRequestUpdateItem.BeforeAttachmentAdd
ms.assetid: b283f394-cb21-9187-a68e-c23d758288fc
ms.date: 06/08/2017
localization_priority: Normal
---


# TaskRequestUpdateItem.BeforeAttachmentAdd event (Outlook)

Occurs before an attachment is added to an instance of the parent object.


## Syntax

_expression_. `BeforeAttachmentAdd`( `_Attachment_` , `_Cancel_` )

_expression_ A variable that represents a [TaskRequestUpdateItem](Outlook.TaskRequestUpdateItem.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Attachment_|Required| **[Attachment](Outlook.Attachment.md)**|The **Attachment** to be added to the item.|
| _Cancel_|Required| **Boolean**|Set to  **True** to cancel the operation; otherwise, set to **False** to allow the **Attachment** to be added.|

## See also


[TaskRequestUpdateItem Object](Outlook.TaskRequestUpdateItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]