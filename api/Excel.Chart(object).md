---
title: Chart object (Excel)
keywords: vbaxl10.chm147072
f1_keywords:
- vbaxl10.chm147072
ms.prod: excel
api_name:
- Excel.Chart
ms.assetid: 179c32ce-49bd-6f36-ea12-89fb5443f3ea
ms.date: 04/16/2019
localization_priority: Priority
---


# Chart object (Excel)

Represents a chart in a workbook.


## Remarks

The chart can be either an embedded chart (contained in a **[ChartObject](Excel.ChartObject.md)** object) or a separate chart sheet.

The **[Charts](Excel.Charts.md)** collection contains a **Chart** object for each chart sheet in a workbook. Use **Charts** (_index_), where _index_ is the chart-sheet index number or name, to return a single **Chart** object. 

The chart _index_ number represents the position of the chart sheet on the workbook tab bar. _Charts(1)_ is the first (leftmost) chart in the workbook; _Charts(Charts.Count)_ is the last (rightmost). 

All chart sheets are included in the index count, even if they are hidden. The chart-sheet name is shown on the workbook tab for the chart. You can use the **[Name](Excel.ChartObject.Name.md)** property of the **ChartObject** object to set or return the chart name. 

The following example changes the color of series 1 on chart sheet 1.

```vb
Charts(1).SeriesCollection(1).Format.Fill.ForeColor.RGB = rgbRed
```

<br/>

The following example moves the chart named Sales to the end of the active workbook.

```vb
Charts("Sales").Move after:=Sheets(Sheets.Count)
```

<br/>

The **Chart** object is also a member of the **[Sheets](Excel.Sheets.md)** collection, which contains all the sheets in the workbook (both chart sheets and worksheets). Use **Sheets** (_index_), where _index_ is the sheet index number or name, to return a single sheet.

When a chart is the active object, you can use the **ActiveChart** property to refer to it. A chart sheet is active if the user has selected it or if it has been activated with the **Activate** method of the **Chart** object or the **[Activate](Excel.ChartObject.Activate.md)** method of the **ChartObject** object. 

The following example activates chart sheet 1, and then sets the chart type and title.

```vb
Charts(1).Activate 
With ActiveChart 
 .Type = xlLine 
 .HasTitle = True 
 .ChartTitle.Text = "January Sales" 
End With
```

<br/>

An embedded chart is active if the user has selected it, or the **ChartObject** object in which it is contained has been activated with the **Activate** method. 

The following example activates embedded chart 1 on worksheet 1 and then sets the chart type and title. Notice that after the embedded chart has been activated, the code in this example is the same as that in the previous example. Using the **ActiveChart** property allows you to write Visual Basic code that can refer to either an embedded chart or a chart sheet (whichever is active).

```vb
Worksheets(1).ChartObjects(1).Activate 
ActiveChart.ChartType = xlLine 
ActiveChart.HasTitle = True 
ActiveChart.ChartTitle.Text = "January Sales"
```

<br/>

When a chart sheet is the active sheet, you can use the **ActiveSheet** property to refer to it. The following example uses the **Activate** method to activate the chart sheet named Chart1, and then sets the interior color for series 1 in the chart to blue.

```vb
Charts("chart1").Activate 
ActiveSheet.SeriesCollection(1).Format.Fill.ForeColor.RGB = rgbBlue
```

## Events

- [Activate](Excel.Chart.Activate(even).md)
- [BeforeDoubleClick](Excel.Chart.BeforeDoubleClick.md)
- [BeforeRightClick](Excel.Chart.BeforeRightClick.md)
- [Calculate](Excel.Chart.Calculate.md)
- [Deactivate](Excel.Chart.Deactivate.md)
- [MouseDown](Excel.Chart.MouseDown.md)
- [MouseMove](Excel.Chart.MouseMove.md)
- [MouseUp](Excel.Chart.MouseUp.md)
- [Resize](Excel.Chart.Resize.md)
- [Select](Excel.Chart.Select(even).md)
- [SeriesChange](Excel.Chart.SeriesChange.md)

## Methods

- [Activate](Excel.Chart.Activate(method).md)
- [ApplyChartTemplate](Excel.Chart.ApplyChartTemplate.md)
- [ApplyDataLabels](Excel.Chart.ApplyDataLabels.md)
- [ApplyLayout](Excel.Chart.ApplyLayout.md)
- [Axes](Excel.Chart.Axes.md)
- [ChartGroups](Excel.Chart.ChartGroups.md)
- [ChartObjects](Excel.Chart.ChartObjects.md)
- [ChartWizard](Excel.Chart.ChartWizard.md)
- [CheckSpelling](Excel.Chart.CheckSpelling.md)
- [ClearToMatchColorStyle](Excel.chart.cleartomatchcolorstyle.md)
- [ClearToMatchStyle](Excel.Chart.ClearToMatchStyle.md)
- [Copy](Excel.Chart.Copy.md)
- [CopyPicture](Excel.Chart.CopyPicture.md)
- [Delete](Excel.Chart.Delete.md)
- [Evaluate](Excel.Chart.Evaluate.md)
- [Export](Excel.Chart.Export.md)
- [ExportAsFixedFormat](Excel.Chart.ExportAsFixedFormat.md)
- [FullSeriesCollection](Excel.chart.fullseriescollection.md)
- [GetChartElement](Excel.Chart.GetChartElement.md)
- [Location](Excel.Chart.Location.md)
- [Move](Excel.Chart.Move.md)
- [OLEObjects](Excel.Chart.OLEObjects.md)
- [Paste](Excel.Chart.Paste.md)
- [PrintOut](Excel.Chart.PrintOut.md)
- [PrintPreview](Excel.Chart.PrintPreview.md)
- [Protect](Excel.Chart.Protect.md)
- [Refresh](Excel.Chart.Refresh.md)
- [SaveAs](Excel.Chart.SaveAs.md)
- [SaveChartTemplate](Excel.Chart.SaveChartTemplate.md)
- [Select](Excel.Chart.Select(method).md)
- [SeriesCollection](Excel.Chart.SeriesCollection.md)
- [SetBackgroundPicture](Excel.Chart.SetBackgroundPicture.md)
- [SetDefaultChart](Excel.Chart.SetDefaultChart.md)
- [SetElement](Excel.Chart.SetElement.md)
- [SetSourceData](Excel.Chart.SetSourceData.md)
- [Unprotect](Excel.Chart.Unprotect.md)

## Properties

- [Application](Excel.Chart.Application.md)
- [AutoScaling](Excel.Chart.AutoScaling.md)
- [BackWall](Excel.Chart.BackWall.md)
- [BarShape](Excel.Chart.BarShape.md)
- [CategoryLabelLevel](Excel.chart.categorylabellevel.md)
- [ChartArea](Excel.Chart.ChartArea.md)
- [ChartColor](Excel.chart.chartcolor.md)
- [ChartStyle](Excel.Chart.ChartStyle.md)
- [ChartTitle](Excel.Chart.ChartTitle.md)
- [ChartType](Excel.Chart.ChartType.md)
- [CodeName](Excel.Chart.CodeName.md)
- [Creator](Excel.Chart.Creator.md)
- [DataTable](Excel.Chart.DataTable.md)
- [DepthPercent](Excel.Chart.DepthPercent.md)
- [DisplayBlanksAs](Excel.Chart.DisplayBlanksAs.md)
- [Elevation](Excel.Chart.Elevation.md)
- [Floor](Excel.Chart.Floor.md)
- [GapDepth](Excel.Chart.GapDepth.md)
- [HasAxis](Excel.Chart.HasAxis.md)
- [HasDataTable](Excel.Chart.HasDataTable.md)
- [HasLegend](Excel.Chart.HasLegend.md)
- [HasTitle](Excel.Chart.HasTitle.md)
- [HeightPercent](Excel.Chart.HeightPercent.md)
- [Hyperlinks](Excel.Chart.Hyperlinks.md)
- [Index](Excel.Chart.Index.md)
- [Legend](Excel.Chart.Legend.md)
- [MailEnvelope](Excel.Chart.MailEnvelope.md)
- [Name](Excel.Chart.Name.md)
- [Next](Excel.Chart.Next.md)
- [PageSetup](Excel.Chart.PageSetup.md)
- [Parent](Excel.Chart.Parent.md)
- [Perspective](Excel.Chart.Perspective.md)
- [PivotLayout](Excel.Chart.PivotLayout.md)
- [PlotArea](Excel.Chart.PlotArea.md)
- [PlotBy](Excel.Chart.PlotBy.md)
- [PlotVisibleOnly](Excel.Chart.PlotVisibleOnly.md)
- [Previous](Excel.Chart.Previous.md)
- [PrintedCommentPages](Excel.Chart.PrintedCommentPages.md)
- [ProtectContents](Excel.Chart.ProtectContents.md)
- [ProtectData](Excel.Chart.ProtectData.md)
- [ProtectDrawingObjects](Excel.Chart.ProtectDrawingObjects.md)
- [ProtectFormatting](Excel.Chart.ProtectFormatting.md)
- [ProtectionMode](Excel.Chart.ProtectionMode.md)
- [ProtectSelection](Excel.Chart.ProtectSelection.md)
- [RightAngleAxes](Excel.Chart.RightAngleAxes.md)
- [Rotation](Excel.Chart.Rotation.md)
- [SeriesNameLevel](Excel.chart.seriesnamelevel.md)
- [Shapes](Excel.Chart.Shapes.md)
- [ShowAllFieldButtons](Excel.Chart.ShowAllFieldButtons.md)
- [ShowAxisFieldButtons](Excel.Chart.ShowAxisFieldButtons.md)
- [ShowDataLabelsOverMaximum](Excel.Chart.ShowDataLabelsOverMaximum.md)
- [ShowExpandCollapseEntireFieldButtons](Excel.chart.showexpandcollapseentirefieldbuttons.md)
- [ShowLegendFieldButtons](Excel.Chart.ShowLegendFieldButtons.md)
- [ShowReportFilterFieldButtons](Excel.Chart.ShowReportFilterFieldButtons.md)
- [ShowValueFieldButtons](Excel.Chart.ShowValueFieldButtons.md)
- [SideWall](Excel.Chart.SideWall.md)
- [Tab](Excel.Chart.Tab.md)
- [Visible](Excel.Chart.Visible.md)
- [Walls](Excel.Chart.Walls.md)

## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
