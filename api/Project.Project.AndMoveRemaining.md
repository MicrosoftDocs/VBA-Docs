---
title: Project.AndMoveRemaining property (Project)
keywords: vbapj.chm131077
f1_keywords:
- vbapj.chm131077
ms.prod: project-server
api_name:
- Project.Project.AndMoveRemaining
ms.assetid: 4ad6b54e-f5b0-b1dc-866f-04ff750300e5
ms.date: 06/08/2017
localization_priority: Normal
---


# Project.AndMoveRemaining property (Project)

 **True** if the remaining work on a task that is scheduled after the status date is moved to start at the status date. Read/write **Boolean**.


## Syntax

_expression_. `AndMoveRemaining`

_expression_ A variable that represents a **[Project](project.project.md)** object.


## Remarks

The **AndMoveCompleted** and **AndMoveRemaining** properties can also be set with the **[OptionsCalculation](Project.Application.OptionsCalculation.md)** method.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]