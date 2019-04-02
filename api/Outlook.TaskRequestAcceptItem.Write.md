---
title: TaskRequestAcceptItem.Write event (Outlook)
ms.prod: outlook
api_name:
- Outlook.TaskRequestAcceptItem.Write
ms.assetid: 005b0f33-1848-101b-2119-cb15eb51f411
ms.date: 06/08/2017
localization_priority: Normal
---


# TaskRequestAcceptItem.Write event (Outlook)

Occurs when an instance of the parent object is saved, either explicitly (for example, using the  **[Save](Outlook.TaskRequestAcceptItem.Save.md)** or **[SaveAs](Outlook.TaskRequestAcceptItem.SaveAs.md)** methods) or implicitly (for example, in response to a prompt when closing the item's inspector).


## Syntax

_expression_. `Write`( `_Cancel_` )

_expression_ A variable that represents a [TaskRequestAcceptItem](Outlook.TaskRequestAcceptItem.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Cancel_|Required| **Boolean**| (Not used in VBScript). **False** when the event occurs. If the event procedure sets this argument to **True**, the save operation is not completed.|

## Remarks

In Microsoft Visual Basic Scripting Edition (VBScript), if you set the return value of this function to  **False**, the save operation is not completed.


## See also


[TaskRequestAcceptItem Object](Outlook.TaskRequestAcceptItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]