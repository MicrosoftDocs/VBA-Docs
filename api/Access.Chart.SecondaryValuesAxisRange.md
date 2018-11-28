---
title: Chart.SecondaryValuesAxisRange property (Access)
keywords: vbaac10.chm6126
f1_keywords:
- vbaac10.chm6126
ms.prod: access
api_name:
- Access.Chart.SecondaryValuesAxisRange
ms.date: 05/02/2018
---


# Chart.SecondaryValuesAxisRange property (Access)

Returns or sets the behavior for representing minimum and maximum values on the secondary values axis. Read/write **[AcAxisRange](Access.AcAxisRange.md)**.


## Syntax

 _expression_ . **SecondaryValuesAxisRange**

 _expression_ A variable that represents a **Chart** object.


## Remarks

**SecondaryValuesAxisMinimum** and **SecondaryValuesAxisMaximum** are enforced when the **SecondaryValuesAxisRange** 
property is set to **Fixed**. Otherwise, the **Auto** setting will determine the range based on the lowest and 
highest values in the set.


## See also

[SecondaryValuesAxisMinimum Property](Access.Chart.SecondaryValuesAxisMinimum.md)

[SecondaryValuesAxisMaximum Property](Access.Chart.SecondaryValuesAxisMaximum.md)

[Chart object](Access.Chart.md)