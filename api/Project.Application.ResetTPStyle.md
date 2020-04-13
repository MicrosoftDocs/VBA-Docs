---
title: Application.ResetTPStyle method (Project)
keywords: vbapj.chm1508
f1_keywords:
- vbapj.chm1508
ms.prod: project-server
api_name:
- Project.Application.ResetTPStyle
ms.assetid: aba4187b-5af3-3a8d-7486-038e9bdae0ae
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.ResetTPStyle method (Project)

Resets the specified Team Planner style to the default values.


## Syntax

_expression_.**ResetTPStyle** (_Style_)

_expression_ An expression that returns an **[Application](Project.Application.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Style_|Required|**PjTeamPlannerStyle**|Can be one of the **[PjTeamPlannerStyle](Project.PjTeamPlannerStyle.md)** constants.|

## Return value

 **Boolean**


## Remarks

The **PjTeamPlannerStyle** constants are equivalent to the five styles shown in the **Format** tab of the **Team Planner Tools** in the ribbon, as follows:


|||
|:-----|:-----|
|**Constant**|**Style**|
|**pjTPActualWork**|**Actual Work**|
|**pjTPLateTask**|**Late Task**|
|**pjTPManualTask**|**Manually Scheduled**|
|**pjTPScheduledWork**|**Auto Scheduled**|
|**pjTPSRA**|**External Task**|

## Example

The following line of code resets the border color and fill color of auto-scheduled assignments in the Team Planner to their default values.


```vb
ResetTPStyle Style:=pjTPScheduledWork
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]