---
title: Application.EditTPStyle method (Project)
keywords: vbapj.chm57
f1_keywords:
- vbapj.chm57
ms.prod: project-server
api_name:
- Project.Application.EditTPStyle
ms.assetid: 71252516-31b5-1184-97f8-da27558620f1
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.EditTPStyle method (Project)

Edits the box and border colors of different types of tasks in the Team Planner view.


## Syntax

_expression_. `EditTPStyle`( `_Style_`, `_FillColor_`, `_BorderColor_` )

 _expression_ An expression that returns an **[Application](Project.Application.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Style_|Required|**PjTeamPlannerStyle**|Can be one of the  **[PjTeamPlannerStyle](Project.PjTeamPlannerStyle.md)** constants, which specify whether the task type is auto scheduled, manually scheduled, actual work, an external task, or a late task.|
| _FillColor_|Optional|**Variant**|Fill color of the specified task type. Can be a hexadecimal RGB value, where red is the last byte.|
| _BorderColor_|Optional|**Variant**|Border color of the specified task type. Can be a hexadecimal RGB value, where red is the last byte.|

## Return value

 **Boolean**


## Remarks

To see the available style colors in the Team Planner view, or to manually format the view, in the  **Team Planner Tools** section of the ribbon, choose the **Format** tab.


## Example

In the following example, the first call to  **EditTPStyle** sets late tasks to medium-dark red with a black border. The second call sets manually scheduled tasks to light red with a gray border.


```vb
Sub ChangeTeamPlannerStyles() 
    EditTPStyle Style:=pjTPLateTask, fillColor:=&H4444FF, bordercolor:=&H0 
    EditTPStyle Style:=pjTPManualTask, fillColor:=&H8888FF, bordercolor:=&H888888 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]