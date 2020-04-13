---
title: Application.ShowIgnoredTaskWarnings method (Project)
keywords: vbapj.chm2178
f1_keywords:
- vbapj.chm2178
ms.prod: project-server
api_name:
- Project.Application.ShowIgnoredTaskWarnings
ms.assetid: 77eeb3ef-511d-af17-56c1-aa717fd7d213
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.ShowIgnoredTaskWarnings method (Project)

Shows any ignored warnings for tasks; turns on the warning symbol in the **Indicators** column.


## Syntax

_expression_. `ShowIgnoredTaskWarnings`

 _expression_ An expression that returns an **[Application](Project.Application.md)** object.


## Return value

 **Boolean**


## Remarks

The **Show Ignored Problems** command is in the **Inspect Task** drop-down menu on the **TASK** ribbon. The **ShowIgnoredTaskWarnings** method sets the **Show warning and suggestion indicators for this task** check box in the **Task Inspector** pane for all tasks.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]