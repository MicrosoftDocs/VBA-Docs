---
title: Application.NewTasksStartOn method (Project)
keywords: vbapj.chm2295
f1_keywords:
- vbapj.chm2295
ms.prod: project-server
api_name:
- Project.Application.NewTasksStartOn
ms.assetid: c5009674-105e-a861-56f0-4847926d6c36
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.NewTasksStartOn method (Project)

Specifies how the start date of a new task is set.


## Syntax

_expression_. `NewTasksStartOn`( `_StartOnDate_` )

 _expression_ An expression that returns an **[Application](Project.Application.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _StartOnDate_|Optional|**PjNewTasksStartOnDate**|Specifies whether new tasks start on the project date, the current date, or no date. Can be one of the  **[PjNewTasksStartOnDate](Project.PjNewTasksStartOnDate.md)** constants. The default is **pjProjectDate**.|

## Return value

 **Boolean**


## Remarks

The  **NewTasksStartOn** method corresponds to the **New tasks created** setting on the **Schedule** tab of the **Project Options** dialog box.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]