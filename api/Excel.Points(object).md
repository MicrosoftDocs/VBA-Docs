---
title: Points object (Excel)
keywords: vbaxl10.chm573072
f1_keywords:
- vbaxl10.chm573072
ms.prod: excel
api_name:
- Excel.Points
ms.assetid: 918dc385-ed61-262e-033f-ba829f5ee8b2
ms.date: 03/30/2019
localization_priority: Normal
---


# Points object (Excel)

A collection of all the **[Point](Excel.Point(object).md)** objects in the specified series in a chart.


## Remarks

Use **[Points](Excel.Series.Points.md)** (_index_), where _index_ is the point index number, to return a single **Point** object. Points are numbered from left to right on the series. `Points(1)` is the leftmost point, and `Points(Points.Count)` is the rightmost point.


## Example

Use the **Points** method of the **Series** object to return the **Points** collection. 

The following example adds a data label to the last point on series one in embedded chart one on worksheet one.

```vb
Dim pts As Points 
Set pts = Worksheets(1).ChartObjects(1).Chart. _ 
 SeriesCollection(1).Points 
pts(pts.Count).ApplyDataLabels type:=xlShowValue
```

<br/>

The following example sets the marker style for the third point in series one in embedded chart one on worksheet one. The specified series must be a 2D line, scatter, or radar series.

```vb
Worksheets(1).ChartObjects(1).Chart. _ 
 SeriesCollection(1).Points(3).MarkerStyle = xlDiamond
```


## Methods

- [Item](Excel.Points.Item.md)

## Properties

- [Application](Excel.Points.Application.md)
- [Count](Excel.Points.Count.md)
- [Creator](Excel.Points.Creator.md)
- [Parent](Excel.Points.Parent.md)


## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
