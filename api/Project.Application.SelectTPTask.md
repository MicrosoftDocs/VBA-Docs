---
title: Application.SelectTPTask method (Project)
keywords: vbapj.chm2192
f1_keywords:
- vbapj.chm2192
ms.prod: project-server
api_name:
- Project.Application.SelectTPTask
ms.assetid: ef27e878-8c80-ad09-157d-f803ec2e7352
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.SelectTPTask method (Project)

Selects the specified task in the Team Planner view.


## Syntax

_expression_. `SelectTPTask`( `_TaskUniqueID_` )

 _expression_ An expression that returns an **[Application](Project.Application.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _TaskUniqueID_|Optional|**Variant**|Unique ID of the task to select.|

## Return value

 **Boolean**


## Remarks

If the Team Planner view is not open, the **SelectTPTask** method generates run-time error 1100, "The method is not available in this situation."


## Example

The following example selects two tasks in the Team Planner view. Task 5 remains selected after task 7 is selected.


```vb
Sub SelectTwoTasks()
    SelectTPTask (5)
    SelectTPTask (7)
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]