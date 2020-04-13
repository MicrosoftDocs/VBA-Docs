---
title: Application.GanttBarStyleLate method (Project)
keywords: vbapj.chm82
f1_keywords:
- vbapj.chm82
ms.prod: project-server
api_name:
- Project.Application.GanttBarStyleLate
ms.assetid: 824760ce-0692-de6a-cf50-90307d94f82a
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.GanttBarStyleLate method (Project)

Shows or hides the late tasks style on the active Gantt chart.


## Syntax

_expression_. `GanttBarStyleLate`( `_Show_` )

 _expression_ An expression that returns an **[Application](Project.Application.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Show_|Optional|**Boolean**|If  **False**, hides the late tasks style. If **True**, shows the late tasks style.|

## Return value

 **Boolean**


## Remarks

On the Ribbon, the **GanttBarStyleLate** method corresponds to the **Late Tasks** checkbox in the **Bar Styles** group on the **Format** tab for **Gantt Chart Tools**.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]