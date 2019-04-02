---
title: TaskItem.Send event (Outlook)
ms.prod: outlook
api_name:
- Outlook.TaskItem.Send
ms.assetid: f634105e-5351-6941-e915-ec63cd703b67
ms.date: 06/08/2017
localization_priority: Normal
---


# TaskItem.Send event (Outlook)

Occurs when the user selects the  **Send** action for an item, or when the **Send** method is called for the item, which is an instance of the parent object.


## Syntax

_expression_. `Send`( `_Cancel_` )

_expression_ A variable that represents a [TaskItem](Outlook.TaskItem.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Cancel_|Required| **Boolean**|(Not used in VBScript).  **False** when the event occurs. If the event procedure sets this argument to **True**, the send operation is not completed and the inspector is left open.|

## Remarks

In Microsoft Visual Basic Scripting Edition (VBScript), if you set the return value of this function to  **False**, the item is not sent.


## See also


[TaskItem Object](Outlook.TaskItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]