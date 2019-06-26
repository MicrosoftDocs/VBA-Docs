---
title: Application.InsertSummaryTask method (Project)
keywords: vbapj.chm2180
f1_keywords:
- vbapj.chm2180
ms.prod: project-server
api_name:
- Project.Application.InsertSummaryTask
ms.assetid: efcbf0d9-5912-d6c4-9204-e939af0193ad
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.InsertSummaryTask method (Project)

Inserts a new summary task above the selected task row or cell in a Gantt chart.


## Syntax

_expression_. `InsertSummaryTask`

 _expression_ An expression that returns an **[Application](Project.Application.md)** object.


## Return value

 **Boolean**


## Remarks

The summary task is the same mode (manual or automatic) as the selected task and is at the level of the selected task. The selected task is indented one level below the new summary task. The  **InsertSummaryTask** method corresponds to the **Summary** command in the **Insert** group of the **TASK** tab on the ribbon.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]