---
title: TaskItem.Forward event (Outlook)
ms.prod: outlook
api_name:
- Outlook.TaskItem.Forward
ms.assetid: 93a74a47-b996-5130-74bb-52a662d58a2b
ms.date: 06/08/2017
localization_priority: Normal
---


# TaskItem.Forward event (Outlook)

Occurs when the user selects the  **Forward** action for an item (which is an instance of the parent object).


## Syntax

_expression_. `Forward`( `_Forward_` , `_Cancel_` )

_expression_ A variable that represents a [TaskItem](Outlook.TaskItem.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Forward_|Required| **Object**|The new item being forwarded.|
| _Cancel_|Required| **Boolean**|(Not used in VBScript).  **False** when the event occurs. If the event procedure sets this argument to **True**, the forward operation is not completed and the new item is not displayed.|

## Remarks

In VBScript, if you set the return value of this function to  **False**, the forward action is not completed and the new item is not displayed.


## See also


[TaskItem Object](Outlook.TaskItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]