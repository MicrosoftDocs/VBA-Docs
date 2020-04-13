---
title: Chart object (Word)
keywords: vbawd10.chm1211
f1_keywords:
- vbawd10.chm1211
ms.prod: word
api_name:
- Word.Chart
ms.assetid: 366a825e-0daf-dbb7-b6f2-e7ce1a5ee2ef
ms.date: 06/08/2017
localization_priority: Normal
---


# Chart object (Word)

Represents a chart in a document.


## Remarks

The Example section describes the following properties and methods for returning a  **Chart** object:




- The **[Chart](Word.InlineShape.Chart.md)** property.
    
- The **[AddChart](overview/Word.md)** method.
    



## Example

The **[InlineShapes](Word.inlineshapes.md)** collection contains an object for each inline shape, including charts, in a document. Use **InlineShapes** (_index_), where Index is the index number of an inline shape, to return a single **InlineShape** object. Use the **[HasChart](Word.InlineShape.HasChart.md)** property to determine whether the **InlineShape** object represents a chart. If the **HasChart** property is set to **True**, use the **[Chart](Word.InlineShape.Chart.md)** property to return a **Chart** object.

You can also use the **[Type](Word.InlineShape.Type.md)** property to determine whether the **InlineShape** object represents a chart. If the **Type** property is set to **WdInlineShapeChart**, the inline shape represents a chart.

The following example verifies whether the first inline shape in the active document represents a chart. If so, the example changes the fore color of the first series on the chart.




```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 .Chart.SeriesCollection(1).Format.Fill.ForeColor.RGB = rgbRed 
 End If 
End With
```

The following example creates a new 3D column chart and adds it to the active document.




```vb
ActiveDocument.InlineShapes.AddChart Type:=xl3DColumn 

```


## Methods



|Name|
|:-----|
|[ApplyChartTemplate](Word.Chart.ApplyChartTemplate.md)|
|[ApplyDataLabels](Word.Chart.ApplyDataLabels.md)|
|[ApplyLayout](Word.Chart.ApplyLayout.md)|
|[Axes](Word.Chart.Axes.md)|
|[ChartWizard](Word.Chart.ChartWizard.md)|
|[ClearToMatchColorStyle](Word.chart.cleartomatchcolorstyle.md)|
|[ClearToMatchStyle](Word.Chart.ClearToMatchStyle.md)|
|[Copy](Word.Chart.Copy.md)|
|[CopyPicture](Word.Chart.CopyPicture.md)|
|[Delete](Word.Chart.Delete.md)|
|[Export](Word.Chart.Export.md)|
|[FullSeriesCollection](Word.chart.fullseriescollection.md)|
|[GetChartElement](Word.Chart.GetChartElement.md)|
|[Paste](Word.Chart.Paste.md)|
|[Refresh](Word.Chart.Refresh.md)|
|[SaveChartTemplate](Word.Chart.SaveChartTemplate.md)|
|[Select](Word.Chart.Select.md)|
|[SeriesCollection](Word.Chart.SeriesCollection.md)|
|[SetBackgroundPicture](Word.Chart.SetBackgroundPicture.md)|
|[SetDefaultChart](Word.Chart.SetDefaultChart.md)|
|[SetElement](Word.Chart.SetElement.md)|
|[SetSourceData](Word.Chart.SetSourceData.md)|

## Properties



|Name|
|:-----|
|[Application](Word.Chart.Application.md)|
|[AutoScaling](Word.Chart.AutoScaling.md)|
|[BackWall](Word.Chart.BackWall.md)|
|[BarShape](Word.Chart.BarShape.md)|
|[CategoryLabelLevel](Word.chart.categorylabellevel.md)|
|[ChartArea](Word.Chart.ChartArea.md)|
|[ChartColor](Word.chart.chartcolor.md)|
|[ChartData](Word.Chart.ChartData.md)|
|[ChartGroups](Word.Chart.ChartGroups.md)|
|[ChartStyle](Word.Chart.ChartStyle.md)|
|[ChartTitle](Word.Chart.ChartTitle.md)|
|[ChartType](Word.Chart.ChartType.md)|
|[Creator](Word.Chart.Creator.md)|
|[DataTable](Word.Chart.DataTable.md)|
|[DepthPercent](Word.Chart.DepthPercent.md)|
|[DisplayBlanksAs](Word.Chart.DisplayBlanksAs.md)|
|[Elevation](Word.Chart.Elevation.md)|
|[Floor](Word.Chart.Floor.md)|
|[GapDepth](Word.Chart.GapDepth.md)|
|[HasAxis](Word.Chart.HasAxis.md)|
|[HasDataTable](Word.Chart.HasDataTable.md)|
|[HasLegend](Word.Chart.HasLegend.md)|
|[HasTitle](Word.Chart.HasTitle.md)|
|[HeightPercent](Word.Chart.HeightPercent.md)|
|[Legend](Word.Chart.Legend.md)|
|[Parent](Word.Chart.Parent.md)|
|[Perspective](Word.Chart.Perspective.md)|
|[PivotLayout](Word.Chart.PivotLayout.md)|
|[PlotArea](Word.Chart.PlotArea.md)|
|[PlotBy](Word.Chart.PlotBy.md)|
|[PlotVisibleOnly](Word.Chart.PlotVisibleOnly.md)|
|[RightAngleAxes](Word.Chart.RightAngleAxes.md)|
|[Rotation](Word.Chart.Rotation.md)|
|[SeriesNameLevel](Word.chart.seriesnamelevel.md)|
|[Shapes](Word.Chart.Shapes.md)|
|[ShowAllFieldButtons](Word.Chart.ShowAllFieldButtons.md)|
|[ShowAxisFieldButtons](Word.Chart.ShowAxisFieldButtons.md)|
|[ShowDataLabelsOverMaximum](Word.Chart.ShowDataLabelsOverMaximum.md)|
|[ShowLegendFieldButtons](Word.Chart.ShowLegendFieldButtons.md)|
|[ShowReportFilterFieldButtons](Word.Chart.ShowReportFilterFieldButtons.md)|
|[ShowValueFieldButtons](Word.Chart.ShowValueFieldButtons.md)|
|[SideWall](Word.Chart.SideWall.md)|
|[Walls](Word.Chart.Walls.md)|

## See also


[Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]