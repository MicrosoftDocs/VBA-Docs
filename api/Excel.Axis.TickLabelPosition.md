---
title: Axis.TickLabelPosition property (Excel)
keywords: vbaxl10.chm561099
f1_keywords:
- vbaxl10.chm561099
ms.prod: excel
api_name:
- Excel.Axis.TickLabelPosition
ms.assetid: 50e27107-6dc5-9097-74f7-331642fb52ac
ms.date: 06/08/2017
localization_priority: Normal
---


# Axis.TickLabelPosition property (Excel)

Describes the position of tick-mark labels on the specified axis. Read/write  **[xlTickLabelPosition](Excel.XlTickLabelPosition.md)**.


## Syntax

_expression_. `TickLabelPosition`

_expression_ A variable that represents an **[Axis](Excel.Axis(object).md)** object.


## Remarks





| **xlTickLabelPosition** can be one of these **xlTickLabelPosition** constants.|
| **xlTickLabelPositionLow**|
| **xlTickLabelPositionNone**|
| **xlTickLabelPositionHigh**|
| **xlTickLabelPositionNextToAxis**|

## Example

This example sets tick-mark labels on the category axis in Chart1 to the high position (above the chart).


```vb
Charts("Chart1").Axes(xlCategory) _ 
 .TickLabelPosition = xlTickLabelPositionHigh
```


## See also


[Axis Object](Excel.Axis(object).md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]