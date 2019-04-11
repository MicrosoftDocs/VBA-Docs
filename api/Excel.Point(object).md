---
title: Point object (Excel)
keywords: vbaxl10.chm575072
f1_keywords:
- vbaxl10.chm575072
ms.prod: excel
api_name:
- Excel.Point
ms.assetid: 48ed9aec-2d29-ec4d-8e55-fca13982c358
ms.date: 03/30/2019
localization_priority: Normal
---


# Point object (Excel)

Represents a single point in a series in a chart.


## Remarks

The **Point** object is a member of the **[Points](Excel.Points(object).md)** collection. The **Points** collection contains all the points in one series.


## Example

Use **[Points](Excel.Series.Points.md)** (_index_), where _index_ is the point index number, to return a single **Point** object. Points are numbered from left to right on the series. `Points(1)` is the leftmost point, and `Points(Points.Count)` is the rightmost point. 

The following example sets the marker style for the third point in series one in embedded chart one on worksheet one. The specified series must be a 2D line, scatter, or radar series.

```vb
Worksheets(1).ChartObjects(1).Chart. _ 
 SeriesCollection(1).Points(3).MarkerStyle = xlDiamond
```


## Methods

- [ApplyDataLabels](Excel.Point.ApplyDataLabels.md)
- [ClearFormats](Excel.Point.ClearFormats.md)
- [Copy](Excel.Point.Copy.md)
- [Delete](Excel.Point.Delete.md)
- [Paste](Excel.Point.Paste.md)
- [PieSliceLocation](Excel.Point.PieSliceLocation.md)
- [Select](Excel.Point.Select.md)

## Properties

- [Application](Excel.Point.Application.md)
- [ApplyPictToEnd](Excel.Point.ApplyPictToEnd.md)
- [ApplyPictToFront](Excel.Point.ApplyPictToFront.md)
- [ApplyPictToSides](Excel.Point.ApplyPictToSides.md)
- [Creator](Excel.Point.Creator.md)
- [DataLabel](Excel.Point.DataLabel.md)
- [Explosion](Excel.Point.Explosion.md)
- [Format](Excel.Point.Format.md)
- [Has3DEffect](Excel.Point.Has3DEffect.md)
- [HasDataLabel](Excel.Point.HasDataLabel.md)
- [Height](Excel.Point.Height.md)
- [InvertIfNegative](Excel.Point.InvertIfNegative.md)
- [IsTotal](Excel.point.istotal.md)
- [Left](Excel.Point.Left.md)
- [MarkerBackgroundColor](Excel.Point.MarkerBackgroundColor.md)
- [MarkerBackgroundColorIndex](Excel.Point.MarkerBackgroundColorIndex.md)
- [MarkerForegroundColor](Excel.Point.MarkerForegroundColor.md)
- [MarkerForegroundColorIndex](Excel.Point.MarkerForegroundColorIndex.md)
- [MarkerSize](Excel.Point.MarkerSize.md)
- [MarkerStyle](Excel.Point.MarkerStyle.md)
- [Name](Excel.Point.Name.md)
- [Parent](Excel.Point.Parent.md)
- [PictureType](Excel.Point.PictureType.md)
- [PictureUnit2](Excel.Point.PictureUnit2.md)
- [SecondaryPlot](Excel.Point.SecondaryPlot.md)
- [Shadow](Excel.Point.Shadow.md)
- [Top](Excel.Point.Top.md)
- [Width](Excel.Point.Width.md)


## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
