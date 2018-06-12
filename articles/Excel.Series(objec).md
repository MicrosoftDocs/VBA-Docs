---
title: Series Object (Excel)
keywords: vbaxl10.chm577072
f1_keywords:
- vbaxl10.chm577072
ms.prod: excel
api_name:
- Excel.Series
ms.assetid: c7d34b32-8172-f7a0-0a17-f01d44246b64
ms.date: 06/08/2017
---


# Series Object (Excel)

Represents a series in a chart.


## Remarks

 The **Series** object is a member of the **[SeriesCollection](Excel.SeriesCollection(object).md)** collection.


## Example

Use  **SeriesCollection** ( _index_ ), where _index_ is the series index number or name, to return a single **Series** object. The following example sets the color of the interior for the first series in embedded chart one on Sheet1.

The series index number indicates the order in which the series were added to the chart.  `SeriesCollection(1)` is the first series added to the chart, and `SeriesCollection(SeriesCollection.Count)` is the last one added.




```
Worksheets("sheet1").ChartObjects(1).Chart. _ 
 SeriesCollection(1).Interior.Color = RGB(255, 0, 0)
```


## Methods



|**Name**|
|:-----|
|[ApplyDataLabels](Excel.Series.ApplyDataLabels.md)|
|[ClearFormats](Excel.Series.ClearFormats.md)|
|[Copy](Excel.Series.Copy.md)|
|[DataLabels](Excel.Series.DataLabels.md)|
|[Delete](Excel.Series.Delete.md)|
|[ErrorBar](Excel.Series.ErrorBar.md)|
|[Paste](Excel.Series.Paste.md)|
|[Points](Excel.Series.Points.md)|
|[Select](Excel.Series.Select.md)|
|[Trendlines](Excel.Series.Trendlines.md)|

## Properties



|**Name**|
|:-----|
|[Application](Excel.Series.Application.md)|
|[ApplyPictToEnd](Excel.Series.ApplyPictToEnd.md)|
|[ApplyPictToFront](Excel.Series.ApplyPictToFront.md)|
|[ApplyPictToSides](Excel.Series.ApplyPictToSides.md)|
|[AxisGroup](Excel.Series.AxisGroup.md)|
|[BarShape](Excel.Series.BarShape.md)|
|[BubbleSizes](Excel.Series.BubbleSizes.md)|
|[ChartType](Excel.Series.ChartType.md)|
|[Creator](Excel.Series.Creator.md)|
|[ErrorBars](Excel.Series.ErrorBars.md)|
|[Explosion](Excel.Series.Explosion.md)|
|[Format](Excel.Series.Format.md)|
|[Formula](Excel.Series.Formula.md)|
|[FormulaLocal](Excel.Series.FormulaLocal.md)|
|[FormulaR1C1](Excel.Series.FormulaR1C1.md)|
|[FormulaR1C1Local](Excel.Series.FormulaR1C1Local.md)|
|[Has3DEffect](Excel.Series.Has3DEffect.md)|
|[HasDataLabels](Excel.Series.HasDataLabels.md)|
|[HasErrorBars](Excel.Series.HasErrorBars.md)|
|[HasLeaderLines](Excel.Series.HasLeaderLines.md)|
|[InvertColor](Excel.Series.InvertColor.md)|
|[InvertColorIndex](Excel.Series.InvertColorIndex.md)|
|[InvertIfNegative](Excel.Series.InvertIfNegative.md)|
|[IsFiltered](Excel.series.isfiltered.md)|
|[LeaderLines](Excel.Series.LeaderLines.md)|
|[MarkerBackgroundColor](Excel.Series.MarkerBackgroundColor.md)|
|[MarkerBackgroundColorIndex](Excel.Series.MarkerBackgroundColorIndex.md)|
|[MarkerForegroundColor](Excel.Series.MarkerForegroundColor.md)|
|[MarkerForegroundColorIndex](Excel.Series.MarkerForegroundColorIndex.md)|
|[MarkerSize](Excel.Series.MarkerSize.md)|
|[MarkerStyle](Excel.Series.MarkerStyle.md)|
|[Name](Excel.Series.Name.md)|
|[Parent](Excel.Series.Parent.md)|
|[PictureType](Excel.Series.PictureType.md)|
|[PictureUnit2](Excel.Series.PictureUnit2.md)|
|[PlotColorIndex](Excel.Series.PlotColorIndex.md)|
|[PlotOrder](Excel.Series.PlotOrder.md)|
|[Shadow](Excel.Series.Shadow.md)|
|[Smooth](Excel.Series.Smooth.md)|
|[Type](Excel.Series.Type.md)|
|[Values](Excel.Series.Values.md)|
|[XValues](Excel.Series.XValues.md)|
|[ParentDataLabelOption](Excel.series.parentdatalabeloption.md)|
|[QuartileCalculationInclusiveMedian](Excel.series.quartilecalculationinclusivemedian.md)|

## See also


#### Other resources


[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)
