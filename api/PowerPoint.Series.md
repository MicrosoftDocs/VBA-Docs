---
title: Series object (PowerPoint)
keywords: vbapp10.chm716000
f1_keywords:
- vbapp10.chm716000
ms.prod: powerpoint
api_name:
- PowerPoint.Series
ms.assetid: 5c8c2d92-d8ca-4d21-e213-c374292275d4
ms.date: 06/08/2017
localization_priority: Normal
---


# Series object (PowerPoint)

Represents a series in a chart.


## Remarks

 The **Series** object is a member of the **[SeriesCollection](PowerPoint.SeriesCollection.md)** collection.


## Example




> [!NOTE] 
> Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

Use  **[SeriesCollection](PowerPoint.Chart.SeriesCollection.md)** (_index_), where _index_ is the series index number or name, to return a single **Series** object. The following example sets the color of the interior for the first series of the first chart in the active document.

The series index number indicates the order in which the series were added to the chart.  `SeriesCollection(1)` is the first series added to the chart, and `SeriesCollection(SeriesCollection.Count)` is the last one added.




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        .Chart.SeriesCollection(1).Interior.Color = RGB(255, 0, 0)

    End If

End With
```


## Methods



|Name|
|:-----|
|[ApplyDataLabels](PowerPoint.Series.ApplyDataLabels.md)|
|[ClearFormats](PowerPoint.Series.ClearFormats.md)|
|[Copy](PowerPoint.Series.Copy.md)|
|[DataLabels](PowerPoint.Series.DataLabels.md)|
|[Delete](PowerPoint.Series.Delete.md)|
|[ErrorBar](PowerPoint.Series.ErrorBar.md)|
|[Paste](PowerPoint.Series.Paste.md)|
|[Points](PowerPoint.Series.Points.md)|
|[Select](PowerPoint.Series.Select.md)|
|[Trendlines](PowerPoint.Series.Trendlines.md)|

## Properties



|Name|
|:-----|
|[Application](PowerPoint.Series.Application.md)|
|[ApplyPictToEnd](PowerPoint.Series.ApplyPictToEnd.md)|
|[ApplyPictToFront](PowerPoint.Series.ApplyPictToFront.md)|
|[ApplyPictToSides](PowerPoint.Series.ApplyPictToSides.md)|
|[AxisGroup](PowerPoint.Series.AxisGroup.md)|
|[BarShape](PowerPoint.Series.BarShape.md)|
|[BubbleSizes](PowerPoint.Series.BubbleSizes.md)|
|[ChartType](PowerPoint.Series.ChartType.md)|
|[Creator](PowerPoint.Series.Creator.md)|
|[ErrorBars](PowerPoint.Series.ErrorBars.md)|
|[Explosion](PowerPoint.Series.Explosion.md)|
|[Format](PowerPoint.Series.Format.md)|
|[Formula](PowerPoint.Series.Formula.md)|
|[FormulaLocal](PowerPoint.Series.FormulaLocal.md)|
|[FormulaR1C1](PowerPoint.Series.FormulaR1C1.md)|
|[FormulaR1C1Local](PowerPoint.Series.FormulaR1C1Local.md)|
|[Has3DEffect](PowerPoint.Series.Has3DEffect.md)|
|[HasDataLabels](PowerPoint.Series.HasDataLabels.md)|
|[HasErrorBars](PowerPoint.Series.HasErrorBars.md)|
|[HasLeaderLines](PowerPoint.Series.HasLeaderLines.md)|
|[InvertColor](PowerPoint.Series.InvertColor.md)|
|[InvertColorIndex](PowerPoint.Series.InvertColorIndex.md)|
|[InvertIfNegative](PowerPoint.Series.InvertIfNegative.md)|
|[IsFiltered](PowerPoint.series.isfiltered.md)|
|[LeaderLines](PowerPoint.Series.LeaderLines.md)|
|[MarkerBackgroundColor](PowerPoint.Series.MarkerBackgroundColor.md)|
|[MarkerBackgroundColorIndex](PowerPoint.Series.MarkerBackgroundColorIndex.md)|
|[MarkerForegroundColor](PowerPoint.Series.MarkerForegroundColor.md)|
|[MarkerForegroundColorIndex](PowerPoint.Series.MarkerForegroundColorIndex.md)|
|[MarkerSize](PowerPoint.Series.MarkerSize.md)|
|[MarkerStyle](PowerPoint.Series.MarkerStyle.md)|
|[Name](PowerPoint.Series.Name.md)|
|[Parent](PowerPoint.Series.Parent.md)|
|**[ParentDataLabelOption](PowerPoint.series.parentdatalabeloption.md)**|
|:-----|
|[PictureType](PowerPoint.Series.PictureType.md)|
|[PictureUnit2](PowerPoint.Series.PictureUnit2.md)|
|[PlotColorIndex](PowerPoint.Series.PlotColorIndex.md)|
|[PlotOrder](PowerPoint.Series.PlotOrder.md)|
|[QuartileCalculationInclusiveMedian](PowerPoint.series.quartilecalculationinclusivemedian.md)|
|[Shadow](PowerPoint.Series.Shadow.md)|
|[Smooth](PowerPoint.Series.Smooth.md)|
|[Type](PowerPoint.Series.Type.md)|
|[Values](PowerPoint.Series.Values.md)|
|[XValues](PowerPoint.Series.XValues.md)|

## See also


[PowerPoint Object Model Reference](overview/PowerPoint/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]