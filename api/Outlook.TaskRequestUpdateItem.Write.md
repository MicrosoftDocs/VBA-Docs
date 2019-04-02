---
title: TaskRequestUpdateItem.Write event (Outlook)
ms.prod: outlook
api_name:
- Outlook.TaskRequestUpdateItem.Write
ms.assetid: afad6071-f421-fc9f-c2b9-d090d5301f35
ms.date: 06/08/2017
localization_priority: Normal
---


# TaskRequestUpdateItem.Write event (Outlook)

Occurs when an instance of the parent object is saved, either explicitly (for example, using the  **[Save](Outlook.TaskRequestUpdateItem.Save.md)** or **[SaveAs](Outlook.TaskRequestUpdateItem.SaveAs.md)** methods) or implicitly (for example, in response to a prompt when closing the item's inspector).


## Syntax

_expression_. `Write`( `_Cancel_` )

_expression_ A variable that represents a [TaskRequestUpdateItem](Outlook.TaskRequestUpdateItem.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Cancel_|Required| **Boolean**| (Not used in VBScript). **False** when the event occurs. If the event procedure sets this argument to **True**, the save operation is not completed.|

## Remarks

In Microsoft Visual Basic Scripting Edition (VBScript), if you set the return value of this function to  **False**, the save operation is not completed.


## See also


[TaskRequestUpdateItem Object](Outlook.TaskRequestUpdateItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]