---
title: Application.GanttBarStyleSlack method (Project)
keywords: vbapj.chm81
f1_keywords:
- vbapj.chm81
ms.prod: project-server
api_name:
- Project.Application.GanttBarStyleSlack
ms.assetid: ccd8feb0-8551-c3fd-3ce5-ca90baaff910
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.GanttBarStyleSlack method (Project)

Shows or hides the slack style for tasks on the active Gantt chart.


## Syntax

_expression_. `GanttBarStyleSlack`( `_Show_` )

 _expression_ An expression that returns an **[Application](Project.Application.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Show_|Optional|**Boolean**|If  **False**, hides the slack. If **True**, shows the slack.|

## Return value

 **Boolean**


## Remarks

On the Ribbon, the **GanttBarStyleSlack** method corresponds to the **Slack** checkbox in the **Bar Styles** group on the **Format** tab for **Gantt Chart Tools**.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]