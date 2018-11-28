---
title: Chart.SecondaryValuesAxisMaximum property (Access)
keywords: vbaac10.chm6122
f1_keywords:
- vbaac10.chm6122
ms.prod: access
api_name:
- Access.Chart.SecondaryValuesAxisMaximum
ms.date: 05/02/2018
---


# Chart.SecondaryValuesAxisMaximum property (Access)

Returns or sets the maximum value that can be represented on the secondary values axis. Read/write **Single**.


## Syntax

 _expression_ . **SecondaryValuesAxisMaximum**

 _expression_ A variable that represents a **Chart** object.


## Remarks

**SecondaryValuesAxisMinimum** and **SecondaryValuesAxisMaximum** are enforced when the **SecondaryValuesAxisRange** 
property is set to **Fixed**.

A chart value may exceed the **SecondaryValuesAxisMaximum** but its representation in a chart (e.g. a bar in a 
bar chart) may be clipped according to the maximum.


## See also


#### Concepts


[SecondaryValuesAxisMinimum Property](Access.Chart.SecondaryValuesAxisMinimum.md)

[SecondaryValuesAxisRange Property](Access.Chart.SecondaryValuesAxisRange.md)

[Chart object](Access.Chart.md)