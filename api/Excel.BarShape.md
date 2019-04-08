---
title: BarShape property (Excel Graph)
keywords: vbagr10.chm66939
f1_keywords:
- vbagr10.chm66939
ms.prod: excel
api_name:
- Excel.BarShape
ms.assetid: 2da9b9aa-84db-6ade-845e-abcb142acc3b
ms.date: 06/08/2017
localization_priority: Normal
---


# BarShape property (Excel Graph)

Returns or sets the shape used with the specified 3-D bar or column chart. Read/write XlBarShape .



|XlBarShape can be one of these XlBarShape constants.|
| **xlConeToMax**|
| **xlCylinder**|
| **xlPyramidToPoint**|
| **xlBox**|
| **xlConeToPoint**|
| **xlPyramidToMax**|

## Syntax

_expression_. `BarShape`

_expression_ Required. An expression that returns one of the objects in the **Applies To** list.

## Example

This example sets the shape used with series one on the chart.


```vb
myChart.SeriesCollection(1).BarShape = xlConeToPoint
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]