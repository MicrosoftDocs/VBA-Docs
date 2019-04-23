---
title: Point object (Excel Graph)
keywords: vbagr10.chm131114
f1_keywords:
- vbagr10.chm131114
ms.prod: excel
api_name:
- Excel.Point
ms.assetid: 944d5edb-b1e7-7aed-5ead-bde3878b26e5
ms.date: 04/06/2019
localization_priority: Normal
---


# Point object (Excel Graph)

Represents a single point in a series on the specified chart. The **Point** object is a member of the **[Points](Excel.points(collection).md)** collection, which contains all the points in the specified series.


## Remarks

Use **Points** (_index_), where _index_ is the point's index number, to return a single **Point** object. 

Points are numbered from left to right in the series. `Points(1)` is the leftmost point, and `Points(Points.Count)` is the rightmost point. 


## Example

The following example sets the marker style for the third point in series one. For this example to work, series one must be a 2D line, scatter, or radar series.

```vb
myChart.SeriesCollection(1).Points(3).MarkerStyle = xlDiamond
```

## See also

- [Excel Graph Visual Basic Reference](overview/excel/graph-visual-basic-reference.md)
- [Excel Object Model Reference](overview/excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]