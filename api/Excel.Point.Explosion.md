---
title: Point.Explosion property (Excel)
keywords: vbaxl10.chm576080
f1_keywords:
- vbaxl10.chm576080
api_name:
- Excel.Point.Explosion
ms.assetid: b6b557c3-d41b-d496-4093-336ec07fb575
ms.date: 05/09/2019
ms.localizationpriority: medium
---


# Point.Explosion property (Excel)

Returns or sets the explosion value for a pie-chart or doughnut-chart slice. Returns 0 (zero) if there's no explosion (the tip of the slice is in the center of the pie). Read/write **Long**.


## Syntax

_expression_.**Explosion**

_expression_ A variable that represents a **[Point](Excel.Point(object).md)** object.


## Example

This example sets the explosion value for point two on Chart1. The example should be run on a pie chart.

```vb
Charts("Chart1").SeriesCollection(1).Points(2).Explosion = 20
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]