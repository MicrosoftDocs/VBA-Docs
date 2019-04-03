---
title: TaskRequestItem.CustomAction event (Outlook)
ms.prod: outlook
api_name:
- Outlook.TaskRequestItem.CustomAction
ms.assetid: 1b4fbc87-6494-b85e-f5a6-c2a538a21078
ms.date: 06/08/2017
localization_priority: Normal
---


# TaskRequestItem.CustomAction event (Outlook)

Occurs when a custom action of an item (which is an instance of the parent object) executes.


## Syntax

_expression_. `CustomAction`( `_Action_` , `_Response_` , `_Cancel_` )

_expression_ A variable that represents a [TaskRequestItem](Outlook.TaskRequestItem.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Action_|Required| **Object**|The  **[Action](Outlook.Action.md)** object.|
| _Response_|Required| **Object**|The newly created item resulting from the custom action.|
| _Cancel_|Required| **Boolean**|(Not used in VBScript).  **False** when the event occurs. If the event procedure sets this argument to **True**, the custom action is not completed.|

## Remarks

The  **Action** object and the newly created item resulting from the custom action are passed to the event.

In Microsoft Visual Basic Scripting Edition (VBScript), if you set the return value of this function to  **False**, the custom action operation is not completed.


## See also


[TaskRequestItem Object](Outlook.TaskRequestItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]