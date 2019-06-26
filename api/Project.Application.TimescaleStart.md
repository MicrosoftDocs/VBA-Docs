---
title: Application.TimescaleStart property (Project)
ms.prod: project-server
api_name:
- Project.Application.TimescaleStart
ms.assetid: 001e0556-e1b4-d817-868a-834970becc46
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.TimescaleStart property (Project)

Gets the date that the timescale in the current view starts. Read-only  **Variant**.


## Syntax

_expression_. `TimescaleStart`

 _expression_ An expression that returns an **[Application](Project.Application.md)** object.


## Remarks

Project adjusts the start of the timescale to the beginning of a time period. To change the timescale duration, use any of the following methods:  **ZoomTimescale**,  **ZoomOut**,  **ZoomIn**, or  **Zoom**.


## Example

If the Gantt chart timescale starts on May 3, 2012, the following statement shows  **5/3/2012** in the **Immediate** pane of the VBE.


```vb
Debug.Print TimescaleStart
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]