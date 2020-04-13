---
title: Assignment.WorkContour property (Project)
keywords: vbapj.chm132828
f1_keywords:
- vbapj.chm132828
ms.prod: project-server
api_name:
- Project.Assignment.WorkContour
ms.assetid: a47a3012-7e5e-febb-d023-368c7c01e065
ms.date: 06/08/2017
localization_priority: Normal
---


# Assignment.WorkContour property (Project)

Gets or sets the type of work contour for the assignment. Read/write  **PjWorkContourType**.


## Syntax

_expression_. `WorkContour`

_expression_ A variable that represents an [Assignment](./Project.Assignment.md) object.


## Remarks

The **WorkContour** property can be one of the following **[PjWorkContourType](Project.PjWorkContourType.md)** constants: **pjBackLoaded**, **pjBell**, **pjContour**, **pjDoublePeak**, **pjEarlyPeak**, **pjFlat**, **pjFrontLoaded**, **pjLatePeak**, or **pjTurtle**. The default value is **pjFlat**.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]