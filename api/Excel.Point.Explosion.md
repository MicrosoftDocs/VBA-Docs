---
title: Point.Explosion property (Excel)
keywords: vbaxl10.chm576080
f1_keywords:
- vbaxl10.chm576080
ms.prod: excel
api_name:
- Excel.Point.Explosion
ms.assetid: b6b557c3-d41b-d496-4093-336ec07fb575
ms.date: 06/08/2017
localization_priority: Normal
---


# Point.Explosion property (Excel)

Returns or sets the explosion value for a pie-chart or doughnut-chart slice. Returns 0 (zero) if there's no explosion (the tip of the slice is in the center of the pie). Read/write  **Long**.


## Syntax

_expression_. `Explosion`

_expression_ A variable that represents a [Point](Excel.Point-graph-object.md) object.


## Example

This example sets the explosion value for point two in Chart1. The example should be run on a pie chart.


```vb
Charts("Chart1").SeriesCollection(1).Points(2).Explosion = 20
```


## See also


[Point Object](Excel.Point(object).md)

