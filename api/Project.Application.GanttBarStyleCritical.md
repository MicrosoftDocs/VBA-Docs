---
title: Application.GanttBarStyleCritical method (Project)
keywords: vbapj.chm80
f1_keywords:
- vbapj.chm80
ms.prod: project-server
api_name:
- Project.Application.GanttBarStyleCritical
ms.assetid: 2db96bf5-2a33-2894-8fcb-dcb4842bba4c
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.GanttBarStyleCritical method (Project)

Shows or hides the critical tasks style on the active Gantt chart.


## Syntax

_expression_. `GanttBarStyleCritical`( `_Show_` )

 _expression_ An expression that returns an **[Application](Project.Application.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Show_|Optional|**Boolean**|If  **False**, hides the critical path style. If **True**, shows the critical task style.|

## Return value

 **Boolean**


## Remarks

On the Ribbon, the  **GanttBarStyleCritical** method corresponds to the **Critical Tasks** check box in the **Bar Styles** group on the **Format** tab for **Gantt Chart Tools**.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]