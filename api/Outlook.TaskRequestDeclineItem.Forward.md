---
title: TaskRequestDeclineItem.Forward event (Outlook)
ms.prod: outlook
api_name:
- Outlook.TaskRequestDeclineItem.Forward
ms.assetid: 22d54abd-a54f-a911-f2f6-b11a721d34b9
ms.date: 06/08/2017
localization_priority: Normal
---


# TaskRequestDeclineItem.Forward event (Outlook)

Occurs when the user selects the  **Forward** action for an item (which is an instance of the parent object).


## Syntax

_expression_. `Forward`( `_Forward_` , `_Cancel_` )

_expression_ A variable that represents a [TaskRequestDeclineItem](Outlook.TaskRequestDeclineItem.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Forward_|Required| **Object**|The new item being forwarded.|
| _Cancel_|Required| **Boolean**|(Not used in VBScript).  **False** when the event occurs. If the event procedure sets this argument to **True**, the forward operation is not completed and the new item is not displayed.|

## Remarks

In VBScript, if you set the return value of this function to  **False**, the forward action is not completed and the new item is not displayed.


## See also


[TaskRequestDeclineItem Object](Outlook.TaskRequestDeclineItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]