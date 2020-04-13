---
title: Application.InsertBlankRow method (Project)
keywords: vbapj.chm2171
f1_keywords:
- vbapj.chm2171
ms.prod: project-server
api_name:
- Project.Application.InsertBlankRow
ms.assetid: 1726e283-d242-53d4-d675-b9cb9d649d29
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.InsertBlankRow method (Project)

Inserts a blank row above the selected task row or cell in a Gantt chart.


## Syntax

_expression_. `InsertBlankRow`

 _expression_ An expression that returns an **[Application](Project.Application.md)** object.


## Return value

 **Boolean**


## Remarks

The **InsertBlankRow** method corresponds to the **Blank Row** command in the **Task** drop-down list in the **Insert** group on the **TASK** tab on the ribbon. The blank row contains only a task ID number, where the empty **Task Name** cell is selected. Each task ID below the new row increases by one. Additional information for the new task can be added programmatically.

By comparison, the **[InsertTask](Project.Application.InsertTask.md)** method creates a task of the default task type, where the **Task Name** cell is selected with **<Type Task Name Here>**.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]