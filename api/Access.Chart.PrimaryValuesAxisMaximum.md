---
title: Chart.PrimaryValuesAxisMaximum property (Access)
keywords: vbaac10.chm6119
f1_keywords:
- vbaac10.chm6119
ms.prod: access
api_name:
- Access.Chart.PrimaryValuesAxisMaximum
ms.date: 11/28/2018
localization_priority: Normal
---


# Chart.PrimaryValuesAxisMaximum property (Access)

Returns or sets the maximum value that can be represented on the primary values axis. Read/write **Single**.


## Syntax

_expression_.**PrimaryValuesAxisMaximum**

_expression_ A variable that represents a **[Chart](Access.Chart.md)** object.


## Remarks

**[PrimaryValuesAxisMinimum](Access.Chart.PrimaryValuesAxisMinimum.md)** and **PrimaryValuesAxisMaximum** are enforced when the **[PrimaryValuesAxisRange](Access.Chart.PrimaryValuesAxisRange.md)** property is set to **Fixed**.

A chart value may exceed the **PrimaryValuesAxisMaximum**, but its representation in a chart (for example, a bar in a bar chart) may be clipped according to the maximum.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]