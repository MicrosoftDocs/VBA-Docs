---
title: TaskRequestDeclineItem.BeforeAutoSave event (Outlook)
ms.prod: outlook
api_name:
- Outlook.TaskRequestDeclineItem.BeforeAutoSave
ms.assetid: a1def448-d1cf-3eae-17c9-aeaafa8fd47b
ms.date: 06/08/2017
localization_priority: Normal
---


# TaskRequestDeclineItem.BeforeAutoSave event (Outlook)

Occurs before the item is automatically saved by Outlook.


## Syntax

_expression_. `BeforeAutoSave`( `_Cancel_` )

_expression_ A variable that represents a [TaskRequestDeclineItem](Outlook.TaskRequestDeclineItem.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Cancel_|Required| **Boolean**|Set to  **True** to cancel the operation; otherwise, set to **False** to allow the **[TaskRequestDeclineItem](Outlook.TaskRequestDeclineItem.md)** to be saved.|

## See also


[TaskRequestDeclineItem Object](Outlook.TaskRequestDeclineItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]