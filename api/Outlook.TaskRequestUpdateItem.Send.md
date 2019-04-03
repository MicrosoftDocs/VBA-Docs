---
title: TaskRequestUpdateItem.Send event (Outlook)
ms.prod: outlook
api_name:
- Outlook.TaskRequestUpdateItem.Send
ms.assetid: 5ae11d3f-67f8-3256-b26f-88a89bade5a1
ms.date: 06/08/2017
localization_priority: Normal
---


# TaskRequestUpdateItem.Send event (Outlook)

Occurs when the user selects the  **Send** action for an item (which is an instance of the parent object).


## Syntax

_expression_. `Send`( `_Cancel_` )

_expression_ A variable that represents a [TaskRequestUpdateItem](Outlook.TaskRequestUpdateItem.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Cancel_|Required| **Boolean**|(Not used in VBScript).  **False** when the event occurs. If the event procedure sets this argument to **True**, the send operation is not completed and the inspector is left open.|

## Remarks

In Microsoft Visual Basic Scripting Edition (VBScript), if you set the return value of this function to  **False**, the item is not sent.


## See also


[TaskRequestUpdateItem Object](Outlook.TaskRequestUpdateItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]