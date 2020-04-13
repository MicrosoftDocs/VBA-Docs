---
title: Application.InsertManualTask method (Project)
keywords: vbapj.chm2169
f1_keywords:
- vbapj.chm2169
ms.prod: project-server
api_name:
- Project.Application.InsertManualTask
ms.assetid: 4fcfa1be-2a92-9906-2024-6bd14a31fdac
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.InsertManualTask method (Project)

Inserts a new manually scheduled task above the selected task row or cell in a Gantt chart.


## Syntax

_expression_. `InsertManualTask`

 _expression_ An expression that returns an **[Application](Project.Application.md)** object.


## Return value

 **Boolean**


## Remarks

By comparison, the **[InsertTask](Project.Application.InsertTask.md)** method creates a task of the default mode and **[InsertScheduledTask](Project.Application.InsertScheduledTask.md)** creates an automatically scheduled task.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]