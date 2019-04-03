---
title: TaskRequestAcceptItem.Forward event (Outlook)
ms.prod: outlook
api_name:
- Outlook.TaskRequestAcceptItem.Forward
ms.assetid: 4437f0b1-0f12-08cf-8661-0e127b5acd3c
ms.date: 06/08/2017
localization_priority: Normal
---


# TaskRequestAcceptItem.Forward event (Outlook)

Occurs when the user selects the  **Forward** action for an item (which is an instance of the parent object).


## Syntax

_expression_. `Forward`( `_Forward_` , `_Cancel_` )

_expression_ A variable that represents a [TaskRequestAcceptItem](Outlook.TaskRequestAcceptItem.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Forward_|Required| **Object**|The new item being forwarded.|
| _Cancel_|Required| **Boolean**|(Not used in VBScript).  **False** when the event occurs. If the event procedure sets this argument to **True**, the forward operation is not completed and the new item is not displayed.|

## Remarks

In VBScript, if you set the return value of this function to  **False**, the forward action is not completed and the new item is not displayed.


## See also


[TaskRequestAcceptItem Object](Outlook.TaskRequestAcceptItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]