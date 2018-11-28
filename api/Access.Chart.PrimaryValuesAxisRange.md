---
title: Chart.PrimaryValuesAxisRange property (Access)
keywords: vbaac10.chm6125
f1_keywords:
- vbaac10.chm6125
ms.prod: access
api_name:
- Access.Chart.PrimaryValuesAxisRange
ms.date: 11/28/2018
---


# Chart.PrimaryValuesAxisRange property (Access)

Returns or sets the behavior for representing minimum and maximum values on the primary values axis. Read/write **[AcAxisRange](Access.AcAxisRange.md)**.


## Syntax

_expression_ . **PrimaryValuesAxisRange**

_expression_ A variable that represents a **[Chart](Access.Chart.md)** object.


## Remarks

**[PrimaryValuesAxisMinimum](Access.Chart.PrimaryValuesAxisMinimum.md)** and **[PrimaryValuesAxisMinimum](Access.Chart.PrimaryValuesAxisMinimum.md)** are enforced when the **PrimaryValuesAxisRange** 
property is set to **Fixed**. Otherwise, the **Auto** setting will determine the range based on the lowest and 
highest values in the set.



