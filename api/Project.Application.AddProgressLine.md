---
title: Application.AddProgressLine method (Project)
keywords: vbapj.chm252
f1_keywords:
- vbapj.chm252
ms.prod: project-server
api_name:
- Project.Application.AddProgressLine
ms.assetid: f7a780f6-63af-e495-9fce-f3f1031bdfa0
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.AddProgressLine method (Project)

Enters interactive progress line mode, enabling the user to manually draw progress lines.


## Syntax

_expression_. `AddProgressLine`

_expression_ A variable that represents an **[Application](Project.Application.md)** object.


## Return value

 **Boolean**


## Remarks

The  **AddProgressLine** method has no effect unless the active view is a Gantt view.

The  **AddProgressLine** method requires user interaction before additional code can be executed.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]