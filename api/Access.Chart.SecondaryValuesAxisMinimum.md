---
title: Chart.SecondaryValuesAxisMinimum property (Access)
keywords: vbaac10.chm6121
f1_keywords:
- vbaac10.chm6121
ms.prod: access
api_name:
- Access.Chart.SecondaryValuesAxisMinimum
ms.date: 11/28/2018
localization_priority: Normal
---


# Chart.SecondaryValuesAxisMinimum property (Access)

Returns or sets the minimum value that can be represented on the secondary values axis. Read/write **Single**.


## Syntax

_expression_.**SecondaryValuesAxisMinimum**

_expression_ A variable that represents a **[Chart](Access.Chart.md)** object.


## Remarks

**SecondaryValuesAxisMinimum** and **[SecondaryValuesAxisMaximum](Access.Chart.SecondaryValuesAxisMaximum.md)** are enforced when the **[SecondaryValuesAxisRange](Access.Chart.SecondaryValuesAxisRange.md)** property is set to **Fixed**.

A chart value may be less than the **SecondaryValuesAxisMinimum**, but its representation in a chart (for example, a bar in a bar chart) may be clipped according to the minimum.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]