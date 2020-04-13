---
title: Application.InsertMilestoneTask method (Project)
keywords: vbapj.chm2170
f1_keywords:
- vbapj.chm2170
ms.prod: project-server
api_name:
- Project.Application.InsertMilestoneTask
ms.assetid: a90ebcc2-b779-0c78-124d-f2c0a9ccd2ca
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.InsertMilestoneTask method (Project)

Inserts a new milestone task above the selected task row or cell in a Gantt chart.


## Syntax

_expression_. `InsertMilestoneTask`

 _expression_ An expression that returns an **[Application](Project.Application.md)** object.


## Return value

 **Boolean**


## Remarks

The milestone task has a duration of 0 days and is of the default mode (automatically or manually scheduled). The **InsertMilestoneTask** method corresponds to the **Milestone** command in the **Insert** group of the **Task** tab on the Ribbon.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]