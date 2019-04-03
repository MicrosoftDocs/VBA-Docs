---
title: TaskRequestAcceptItem.BeforeAttachmentWriteToTempFile event (Outlook)
ms.prod: outlook
api_name:
- Outlook.TaskRequestAcceptItem.BeforeAttachmentWriteToTempFile
ms.assetid: 655b5236-ce22-62b7-0cbc-3bc132ffc992
ms.date: 06/08/2017
localization_priority: Normal
---


# TaskRequestAcceptItem.BeforeAttachmentWriteToTempFile event (Outlook)

Occurs before an attachment associated with an instance of the parent object is written to a temporary file.


## Syntax

_expression_. `BeforeAttachmentWriteToTempFile`( `_Attachment_` , `_Cancel_` )

_expression_ A variable that represents a [TaskRequestAcceptItem](Outlook.TaskRequestAcceptItem.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Attachment_|Required| **[Attachment](Outlook.Attachment.md)**|The  **Attachment** to be written.|
| _Cancel_|Required| **Boolean**|Set to  **True** to cancel the operation; otherwise, set to **False** to allow the **Attachment** to be written.|

## See also


[TaskRequestAcceptItem Object](Outlook.TaskRequestAcceptItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]