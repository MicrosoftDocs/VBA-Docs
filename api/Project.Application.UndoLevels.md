---
title: Application.UndoLevels property (Project)
keywords: vbapj.chm132407
f1_keywords:
- vbapj.chm132407
ms.prod: project-server
api_name:
- Project.Application.UndoLevels
ms.assetid: 2cfd6962-2cae-b7fe-2c8d-f0c81a1c1302
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.UndoLevels property (Project)

Gets or sets the number of undo levels. The default is 20. Read/write  **Long**.


## Syntax

_expression_. `UndoLevels`

_expression_ A variable that represents an **[Application](Project.Application.md)** object.


## Remarks

The minimum value of  **UndoLevels** is 1 (no multilevel undo) and the maximum is 99.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]