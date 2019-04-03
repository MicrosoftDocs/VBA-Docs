---
title: TaskItem.BeforeAutoSave event (Outlook)
ms.prod: outlook
api_name:
- Outlook.TaskItem.BeforeAutoSave
ms.assetid: 390578bf-3c8f-31f1-d81f-e2abba3c1fb6
ms.date: 06/08/2017
localization_priority: Normal
---


# TaskItem.BeforeAutoSave event (Outlook)

Occurs before the item is automatically saved by Outlook.


## Syntax

_expression_. `BeforeAutoSave`( `_Cancel_` , )

_expression_ A variable that represents a [TaskItem](Outlook.TaskItem.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Cancel_|Required| **Boolean**|Set to  **True** to cancel the operation; otherwise, set to **False** to allow the **[TaskItem](Outlook.TaskItem.md)** to be saved.|

## See also


[TaskItem Object](Outlook.TaskItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]