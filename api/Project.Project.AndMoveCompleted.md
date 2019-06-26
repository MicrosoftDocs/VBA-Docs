---
title: Project.AndMoveCompleted property (Project)
keywords: vbapj.chm131076
f1_keywords:
- vbapj.chm131076
ms.prod: project-server
api_name:
- Project.Project.AndMoveCompleted
ms.assetid: 9f14e1e6-0a1e-1a8b-112e-600b3cb46a56
ms.date: 06/08/2017
localization_priority: Normal
---


# Project.AndMoveCompleted property (Project)

 **True** if the actual, completed portion of a task that is scheduled before the status date is moved to end at the status date. Read/write **Boolean**.


## Syntax

_expression_. `AndMoveCompleted`

_expression_ A variable that represents a **[Project](project.project.md)** object.


## Remarks

The  **AndMoveCompleted** and **AndMoveRemaining** properties can also be set with the **[OptionsCalculation](Project.Application.OptionsCalculation.md)** method.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]