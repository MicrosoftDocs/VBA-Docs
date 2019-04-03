---
title: TaskRequestItem.BeforeAutoSave event (Outlook)
ms.prod: outlook
api_name:
- Outlook.TaskRequestItem.BeforeAutoSave
ms.assetid: 0907ec19-5b94-619e-dcd1-8c458294194f
ms.date: 06/08/2017
localization_priority: Normal
---


# TaskRequestItem.BeforeAutoSave event (Outlook)

Occurs before the item is automatically saved by Outlook.


## Syntax

_expression_. `BeforeAutoSave`( `_Cancel_` )

_expression_ A variable that represents a [TaskRequestItem](Outlook.TaskRequestItem.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Cancel_|Required| **Boolean**|Set to  **True** to cancel the operation; otherwise, set to **False** to allow the **[TaskRequestItem](Outlook.TaskRequestItem.md)** to be saved.|

## See also


[TaskRequestItem Object](Outlook.TaskRequestItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]