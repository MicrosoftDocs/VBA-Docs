---
title: TaskRequestAcceptItem.BeforeAutoSave event (Outlook)
ms.prod: outlook
api_name:
- Outlook.TaskRequestAcceptItem.BeforeAutoSave
ms.assetid: 03c76bb7-b267-7c5f-37aa-dd28576b6a65
ms.date: 06/08/2017
localization_priority: Normal
---


# TaskRequestAcceptItem.BeforeAutoSave event (Outlook)

Occurs before the item is automatically saved by Outlook.


## Syntax

_expression_. `BeforeAutoSave`( `_Cancel_` )

_expression_ A variable that represents a [TaskRequestAcceptItem](Outlook.TaskRequestAcceptItem.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Cancel_|Required| **Boolean**|Set to  **True** to cancel the operation; otherwise, set to **False** to allow the **[TaskRequestAcceptItem](Outlook.TaskRequestAcceptItem.md)** to be saved.|

## See also


[TaskRequestAcceptItem Object](Outlook.TaskRequestAcceptItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]