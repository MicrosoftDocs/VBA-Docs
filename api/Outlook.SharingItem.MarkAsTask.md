---
title: SharingItem.MarkAsTask method (Outlook)
keywords: vbaol11.chm3223
f1_keywords:
- vbaol11.chm3223
ms.prod: outlook
api_name:
- Outlook.SharingItem.MarkAsTask
ms.assetid: deab1b6c-2d22-678c-1a13-2b171d27a971
ms.date: 06/08/2017
localization_priority: Normal
---


# SharingItem.MarkAsTask method (Outlook)

Marks a  **[SharingItem](Outlook.SharingItem.md)** object as a task and assigns a task interval for the object.


## Syntax

_expression_. `MarkAsTask`( `_MarkInterval_` )

 _expression_ An expression that returns a [SharingItem](Outlook.SharingItem.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _MarkInterval_|Required| **[OlMarkInterval](Outlook.OlMarkInterval.md)**|The task interval for the  **SharingItem**.|

## Remarks

Calling this method sets the  **[IsMarkedAsTask](Outlook.SharingItem.IsMarkedAsTask.md)** property to **True** and updates the **[TaskStartDate](Outlook.SharingItem.TaskStartDate.md)**, **[TaskDueDate](Outlook.SharingItem.TaskDueDate.md)**, and **[TaskOrdinal](Outlook.SharingItem.ToDoTaskOrdinal.md)** properties depending on the value provided in _MarkInterval_.


## See also


[SharingItem Object](Outlook.SharingItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]