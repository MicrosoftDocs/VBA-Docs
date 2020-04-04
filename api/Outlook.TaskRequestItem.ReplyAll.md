---
title: TaskRequestItem.ReplyAll event (Outlook)
ms.prod: outlook
api_name:
- Outlook.TaskRequestItem.ReplyAll
ms.assetid: f68693c4-9cfd-2805-3559-4117c46f05c8
ms.date: 06/08/2017
localization_priority: Normal
---


# TaskRequestItem.ReplyAll event (Outlook)

Occurs when the user selects the  **ReplyAll** action for an item (which is an instance of the parent object).


## Syntax

_expression_. `ReplyAll`( `_Response_` , `_Cancel_` )

_expression_ A variable that represents a [TaskRequestItem](Outlook.TaskRequestItem.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Response_|Required| **Object**|The new item being sent in response to the original message.|
| _Cancel_|Required| **Boolean**| **False** when the event occurs. If the event procedure sets this argument to **True**, the reply all operation is not completed and the new item is not displayed.|

## Remarks

Returns the reply as a **[MailItem](Outlook.MailItem.md)** object.


## See also


[TaskRequestItem Object](Outlook.TaskRequestItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]