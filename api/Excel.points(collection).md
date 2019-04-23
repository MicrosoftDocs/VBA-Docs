---
title: Points collection (Excel Graph)
keywords: vbagr10.chm5207812
f1_keywords:
- vbagr10.chm5207812
ms.prod: excel
ms.assetid: b41c8f08-880e-1f4a-0456-3f77c0741bc6
ms.date: 04/06/2019
localization_priority: Normal
---


# Points collection (Excel Graph)

A collection of all the **[Point](Excel.Point-graph-object.md)** objects in the specified series in a chart.


## Remarks

Use the **[Points](excel.points-graph-method.md)** method to return the **Points** collection. 

Use **Points** (_index_), where _index_ is the point's index number, to return a single **Point** object. Points are numbered from left to right in the series. `Points(1)` is the leftmost point, and `Points(Points.Count)` is the rightmost point. 


## Example

The following example adds a data label to the last point in series one in the chart.

```vb
Dim pts As Points 
Set pts = myChart.SeriesCollection(1).Points 
pts(pts.Count).ApplyDataLabels Type:=xlShowValue
```

<br/>

The following example sets the marker style for the third point in series one in the chart. The specified series must be a 2D line, scatter, or radar series.

```vb
myChart.SeriesCollection(1).Points(3).MarkerStyle = xlDiamond
```


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]