---
title: Axis.MinorTickMark property (Excel)
keywords: vbaxl10.chm561093
f1_keywords:
- vbaxl10.chm561093
ms.prod: excel
api_name:
- Excel.Axis.MinorTickMark
ms.assetid: 27b0ab28-4690-e493-9eb9-8651bec5ccb8
ms.date: 06/08/2017
localization_priority: Normal
---


# Axis.MinorTickMark property (Excel)

Returns or sets the type of minor tick mark for the specified axis. Read/write  **[xlTickMark](Excel.XlTickMark.md)**.


## Syntax

_expression_. `MinorTickMark`

_expression_ A variable that represents an **[Axis](Excel.Axis(object).md)** object.


## Remarks





| **xlTickMark** can be one of these **xlTickMark** constants.|
| **xlTickMarkInside**|
| **xlTickMarkOutside**|
| **xlTickMarkCross**|
| **xlTickMarkNone**|

## Example

This example sets the minor tick marks for the value axis in Chart1 to be inside the axis.


```vb
Charts("Chart1").Axes(xlValue).MinorTickMark = xlTickMarkInside
```


## See also


[Axis Object](Excel.Axis(object).md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]