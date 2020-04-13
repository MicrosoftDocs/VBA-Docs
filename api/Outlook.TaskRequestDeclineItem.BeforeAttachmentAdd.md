---
title: TaskRequestDeclineItem.BeforeAttachmentAdd event (Outlook)
ms.prod: outlook
api_name:
- Outlook.TaskRequestDeclineItem.BeforeAttachmentAdd
ms.assetid: 1962b6f9-99de-70e7-6845-dc2c547d7035
ms.date: 06/08/2017
localization_priority: Normal
---


# TaskRequestDeclineItem.BeforeAttachmentAdd event (Outlook)

Occurs before an attachment is added to an instance of the parent object.


## Syntax

_expression_. `BeforeAttachmentAdd`( `_Attachment_` , `_Cancel_` )

_expression_ A variable that represents a [TaskRequestDeclineItem](Outlook.TaskRequestDeclineItem.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Attachment_|Required| **[Attachment](Outlook.Attachment.md)**|The **Attachment** to be added to the item.|
| _Cancel_|Required| **Boolean**|Set to  **True** to cancel the operation; otherwise, set to **False** to allow the **Attachment** to be added.|

## See also


[TaskRequestDeclineItem Object](Outlook.TaskRequestDeclineItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]