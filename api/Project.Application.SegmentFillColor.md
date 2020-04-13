---
title: Application.SegmentFillColor method (Project)
keywords: vbapj.chm71
f1_keywords:
- vbapj.chm71
ms.prod: project-server
api_name:
- Project.Application.SegmentFillColor
ms.assetid: 3f943b8a-47e9-979a-4755-f7b021db6b0e
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.SegmentFillColor method (Project)

Sets the fill color for the assignment segments of a selected task in the Team Planner view.


## Syntax

_expression_. `SegmentFillColor`( `_Color_` )

 _expression_ An expression that returns an **[Application](Project.Application.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Color_|Required|**Long**|Fill color of the assignment segments. The color is a hexadecimal RGB value, where red is the last byte.|

## Return value

 **Boolean**


## Example

In the following example, a task is assigned to two resources. After selecting either of the assignments, running the **ChangeSegmentColor** macro shows all assignments for the task as light red with a blue border.


```vb
Sub ChangeSegmentColor() 
    Application.SegmentFillColor(&H8080FF) 
    Application.SegmentBorderColor(&HFF1010) 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]