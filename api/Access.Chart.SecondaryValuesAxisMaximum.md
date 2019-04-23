---
title: Chart.SecondaryValuesAxisMaximum property (Access)
keywords: vbaac10.chm6122
f1_keywords:
- vbaac10.chm6122
ms.prod: access
api_name:
- Access.Chart.SecondaryValuesAxisMaximum
ms.date: 11/28/2018
localization_priority: Normal
---


# Chart.SecondaryValuesAxisMaximum property (Access)

Returns or sets the maximum value that can be represented on the secondary values axis. Read/write **Single**.


## Syntax

_expression_.**SecondaryValuesAxisMaximum**

_expression_ A variable that represents a **[Chart](Access.Chart.md)** object.


## Remarks

**[SecondaryValuesAxisMinimum](Access.Chart.SecondaryValuesAxisMinimum.md)** and **SecondaryValuesAxisMaximum** are enforced when the **[SecondaryValuesAxisRange](Access.Chart.SecondaryValuesAxisRange.md)** property is set to **Fixed**.

A chart value may exceed the **SecondaryValuesAxisMaximum**, but its representation in a chart (for example, a bar in a bar chart) may be clipped according to the maximum.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]