---
title: Chart Object (Word)
keywords: vbawd10.chm1211
f1_keywords:
- vbawd10.chm1211
ms.prod: word
api_name:
- Word.Chart
ms.assetid: 366a825e-0daf-dbb7-b6f2-e7ce1a5ee2ef
ms.date: 06/08/2017
---


# Chart Object (Word)

Represents a chart in a document.


## Remarks

The Example section describes the following properties and methods for returning a  **Chart** object:




- The  **[Chart](../../../api/Word.InlineShape.Chart.md)** property.
    
- The  **[AddChart](http://msdn.microsoft.com/library/1b168e7b-543a-a817-51b0-8171beecc946%28Office.15%29.aspx)** method.
    



## Example

The  **[InlineShapes](../../../api/Word.inlineshapes.md)** collection contains an object for each inline shape, including charts, in a document. Use **InlineShapes** ( _Index_ ), where Index is the index number of an inline shape, to return a single **InlineShape** object. Use the **[HasChart](../../../api/Word.InlineShape.HasChart.md)** property to determine whether the **InlineShape** object represents a chart. If the **HasChart** property is set to **True**, use the **[Chart](../../../api/Word.InlineShape.Chart.md)** property to return a **Chart** object.

You can also use the  **[Type](../../../api/Word.InlineShape.Type.md)** property to determine whether the **InlineShape** object represents a chart. If the **Type** property is set to **WdInlineShapeChart**, the inline shape represents a chart.

The following example verifies whether the first inline shape in the active document represents a chart. If so, the example changes the fore color of the first series on the chart.




```
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 .Chart.SeriesCollection(1).Format.Fill.ForeColor.RGB = rgbRed 
 End If 
End With
```

The following example creates a new 3-D column chart and adds it to the active document.




```
ActiveDocument.InlineShapes.AddChart Type:=xl3DColumn 

```


## Methods



|**Name**|
|:-----|
|[ApplyChartTemplate](../../../api/Word.Chart.ApplyChartTemplate.md)|
|[ApplyDataLabels](../../../api/Word.Chart.ApplyDataLabels.md)|
|[ApplyLayout](../../../api/Word.Chart.ApplyLayout.md)|
|[Axes](../../../api/Word.Chart.Axes.md)|
|[ChartWizard](../../../api/Word.Chart.ChartWizard.md)|
|[ClearToMatchColorStyle](../../../api/Word.chart.cleartomatchcolorstyle.md)|
|[ClearToMatchStyle](../../../api/Word.Chart.ClearToMatchStyle.md)|
|[Copy](../../../api/Word.Chart.Copy.md)|
|[CopyPicture](../../../api/Word.Chart.CopyPicture.md)|
|[Delete](../../../api/Word.Chart.Delete.md)|
|[Export](../../../api/Word.Chart.Export.md)|
|[FullSeriesCollection](../../../api/Word.chart.fullseriescollection.md)|
|[GetChartElement](../../../api/Word.Chart.GetChartElement.md)|
|[Paste](../../../api/Word.Chart.Paste.md)|
|[Refresh](../../../api/Word.Chart.Refresh.md)|
|[SaveChartTemplate](../../../api/Word.Chart.SaveChartTemplate.md)|
|[Select](../../../api/Word.Chart.Select.md)|
|[SeriesCollection](../../../api/Word.Chart.SeriesCollection.md)|
|[SetBackgroundPicture](../../../api/Word.Chart.SetBackgroundPicture.md)|
|[SetDefaultChart](../../../api/Word.Chart.SetDefaultChart.md)|
|[SetElement](../../../api/Word.Chart.SetElement.md)|
|[SetSourceData](chart-setsourcedata-method-word.md)|

## Properties



|**Name**|
|:-----|
|[Application](../../../api/Word.Chart.Application.md)|
|[AutoScaling](../../../api/Word.Chart.AutoScaling.md)|
|[BackWall](../../../api/Word.Chart.BackWall.md)|
|[BarShape](../../../api/Word.Chart.BarShape.md)|
|[CategoryLabelLevel](../../../api/Word.chart.categorylabellevel.md)|
|[ChartArea](chart-chartarea-property-word.md)|
|[ChartColor](../../../api/Word.chart.chartcolor.md)|
|[ChartData](../../../api/Word.Chart.ChartData.md)|
|[ChartGroups](../../../api/Word.Chart.ChartGroups.md)|
|[ChartStyle](../../../api/Word.Chart.ChartStyle.md)|
|[ChartTitle](../../../api/Word.Chart.ChartTitle.md)|
|[ChartType](../../../api/Word.Chart.ChartType.md)|
|[Creator](../../../api/Word.Chart.Creator.md)|
|[DataTable](../../../api/Word.Chart.DataTable.md)|
|[DepthPercent](../../../api/Word.Chart.DepthPercent.md)|
|[DisplayBlanksAs](../../../api/Word.Chart.DisplayBlanksAs.md)|
|[Elevation](../../../api/Word.Chart.Elevation.md)|
|[Floor](../../../api/Word.Chart.Floor.md)|
|[GapDepth](../../../api/Word.Chart.GapDepth.md)|
|[HasAxis](../../../api/Word.Chart.HasAxis.md)|
|[HasDataTable](../../../api/Word.Chart.HasDataTable.md)|
|[HasLegend](../../../api/Word.Chart.HasLegend.md)|
|[HasTitle](../../../api/Word.Chart.HasTitle.md)|
|[HeightPercent](../../../api/Word.Chart.HeightPercent.md)|
|[Legend](../../../api/Word.Chart.Legend.md)|
|[Parent](../../../api/Word.Chart.Parent.md)|
|[Perspective](../../../api/Word.Chart.Perspective.md)|
|[PivotLayout](../../../api/Word.Chart.PivotLayout.md)|
|[PlotArea](../../../api/Word.Chart.PlotArea.md)|
|[PlotBy](../../../api/Word.Chart.PlotBy.md)|
|[PlotVisibleOnly](../../../api/Word.Chart.PlotVisibleOnly.md)|
|[RightAngleAxes](../../../api/Word.Chart.RightAngleAxes.md)|
|[Rotation](../../../api/Word.Chart.Rotation.md)|
|[SeriesNameLevel](../../../api/Word.chart.seriesnamelevel.md)|
|[Shapes](../../../api/Word.Chart.Shapes.md)|
|[ShowAllFieldButtons](../../../api/Word.Chart.ShowAllFieldButtons.md)|
|[ShowAxisFieldButtons](../../../api/Word.Chart.ShowAxisFieldButtons.md)|
|[ShowDataLabelsOverMaximum](../../../api/Word.Chart.ShowDataLabelsOverMaximum.md)|
|[ShowLegendFieldButtons](../../../api/Word.Chart.ShowLegendFieldButtons.md)|
|[ShowReportFilterFieldButtons](../../../api/Word.Chart.ShowReportFilterFieldButtons.md)|
|[ShowValueFieldButtons](../../../api/Word.Chart.ShowValueFieldButtons.md)|
|[SideWall](../../../api/Word.Chart.SideWall.md)|
|[Walls](chart-walls-property-word.md)|

## See also


#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)
