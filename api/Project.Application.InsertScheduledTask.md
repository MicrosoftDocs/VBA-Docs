---
title: Application.InsertScheduledTask method (Project)
keywords: vbapj.chm2168
f1_keywords:
- vbapj.chm2168
ms.prod: project-server
api_name:
- Project.Application.InsertScheduledTask
ms.assetid: 0bf89c86-6e0b-19fb-131c-70be563876bd
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.InsertScheduledTask method (Project)

Inserts a new automatically scheduled task above the selected task row or cell in a Gantt chart.


## Syntax

_expression_. `InsertScheduledTask`

 _expression_ An expression that returns an **[Application](Project.Application.md)** object.


## Return value

 **Boolean**


## Remarks

By comparison, the  **[InsertTask](Project.Application.InsertTask.md)** method creates a task of the default mode and **[InsertManualTask](Project.Application.InsertManualTask.md)** creates a manually scheduled task.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]