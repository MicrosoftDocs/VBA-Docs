---
title: TaskRequestItem.Send event (Outlook)
ms.prod: outlook
api_name:
- Outlook.TaskRequestItem.Send
ms.assetid: 0d47820a-40c6-b975-18c0-1adcc7122bca
ms.date: 06/08/2017
localization_priority: Normal
---


# TaskRequestItem.Send event (Outlook)

Occurs when the user selects the  **Send** action for an item (which is an instance of the parent object).


## Syntax

_expression_. `Send`( `_Cancel_` )

_expression_ A variable that represents a [TaskRequestItem](Outlook.TaskRequestItem.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Cancel_|Required| **Boolean**|(Not used in VBScript).  **False** when the event occurs. If the event procedure sets this argument to **True**, the send operation is not completed and the inspector is left open.|

## Remarks

In Microsoft Visual Basic Scripting Edition (VBScript), if you set the return value of this function to  **False**, the item is not sent.


## See also


[TaskRequestItem Object](Outlook.TaskRequestItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]