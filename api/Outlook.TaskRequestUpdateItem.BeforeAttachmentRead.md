---
title: TaskRequestUpdateItem.BeforeAttachmentRead event (Outlook)
ms.prod: outlook
api_name:
- Outlook.TaskRequestUpdateItem.BeforeAttachmentRead
ms.assetid: 74e4e5d6-d70a-4d1f-1331-18a40b17760d
ms.date: 06/08/2017
localization_priority: Normal
---


# TaskRequestUpdateItem.BeforeAttachmentRead event (Outlook)

Occurs before an attachment associated with an instance of the parent object is read from the file system, an attachment stream, or an **[Attachment](Outlook.Attachment.md)** object.


## Syntax

_expression_. `BeforeAttachmentRead`( `_Attachment_` , `_Cancel_` )

_expression_ A variable that represents a [TaskRequestUpdateItem](Outlook.TaskRequestUpdateItem.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Attachment_|Required| **Attachment**|The **Attachment** to be read.|
| _Cancel_|Required| **Boolean**|Set to  **True** to cancel the operation; otherwise, set to **False** to allow the **Attachment** to be read.|

## See also


[TaskRequestUpdateItem Object](Outlook.TaskRequestUpdateItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]