---
title: Chart.PrimaryValuesAxisMinimum property (Access)
keywords: vbaac10.chm6120
f1_keywords:
- vbaac10.chm6120
ms.prod: access
api_name:
- Access.Chart.PrimaryValuesAxisMinimum
ms.date: 11/28/2018
localization_priority: Normal
---


# Chart.PrimaryValuesAxisMinimum property (Access)

Returns or sets the minimum value that can be represented on the primary values axis. Read/write **Single**.


## Syntax

_expression_.**PrimaryValuesAxisMinimum**

_expression_ A variable that represents a **[Chart](Access.Chart.md)** object.


## Remarks

**PrimaryValuesAxisMinimum** and **[PrimaryValuesAxisMaximum](Access.Chart.PrimaryValuesAxisMaximum.md)** are enforced when the **[PrimaryValuesAxisRange](Access.Chart.PrimaryValuesAxisRange.md)** property is set to **Fixed**.

A chart value may be less than the **PrimaryValuesAxisMinimum**, but its representation in a chart (for example, a bar in a bar chart) may be clipped according to the minimum.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]