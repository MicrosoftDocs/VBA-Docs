---
title: TaskItem.Write event (Outlook)
ms.prod: outlook
api_name:
- Outlook.TaskItem.Write
ms.assetid: 29e38bc5-6a19-5144-55ba-207215bd5734
ms.date: 06/08/2017
localization_priority: Normal
---


# TaskItem.Write event (Outlook)

Occurs when an instance of the parent object is saved, either explicitly (for example, using the  **[Save](Outlook.TaskItem.Save.md)** or **[SaveAs](Outlook.TaskItem.SaveAs.md)** methods) or implicitly (for example, in response to a prompt when closing the item's inspector).


## Syntax

_expression_. `Write`( `_Cancel_` )

_expression_ A variable that represents a [TaskItem](Outlook.TaskItem.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Cancel_|Required| **Boolean**| (Not used in VBScript). **False** when the event occurs. If the event procedure sets this argument to **True**, the save operation is not completed.|

## Remarks

In Microsoft Visual Basic Scripting Edition (VBScript), if you set the return value of this function to  **False**, the save operation is not completed.


## See also


[TaskItem Object](Outlook.TaskItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]