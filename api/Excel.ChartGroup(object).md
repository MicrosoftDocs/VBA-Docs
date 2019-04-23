---
title: ChartGroup object (Excel)
keywords: vbaxl10.chm567072
f1_keywords:
- vbaxl10.chm567072
ms.prod: excel
api_name:
- Excel.ChartGroup
ms.assetid: 7eee66c5-04a7-fd86-6e34-4c22ccaf8de0
ms.date: 03/29/2019
localization_priority: Normal
---


# ChartGroup object (Excel)

Represents one or more series plotted in a chart with the same format.


## Remarks

A chart contains one or more chart groups, each chart group contains one or more **[Series](Excel.Series(object).md)** objects, and each series contains one or more **[Points](Excel.Point(object).md)** objects. 

For example, a single chart might contain both a line chart group that contains all the series plotted with the line chart format, and a bar chart group that contains all the series plotted with the bar chart format. The **ChartGroup** object is a member of the **[ChartGroups](Excel.ChartGroups(object).md)** collection.

Use **ChartGroups** (_index_), where _index_ is the chart-group index number, to return a single **ChartGroup** object.

Because the index number for a particular chart group can change if the chart format used for that group is changed, it may be easier to use one of the named chart group shortcut methods to return a particular chart group. The **[PieGroups](excel.piegroups.md)** method returns the collection of pie chart groups in a chart, the **[LineGroups](excel.linegroups.md)** method returns the collection of line chart groups, and so on. Each of these methods can be used with an index number to return a single **ChartGroup** object, or without an index number to return a **ChartGroups** collection.


## Example

The following example adds drop lines to chart group 1 on chart sheet 1.

```vb
Charts(1).ChartGroups(1).HasDropLines = True
```

<br/>

If the chart has been activated, you can use the **ActiveChart** property.

```vb
Charts(1).Activate 
ActiveChart.ChartGroups(1).HasDropLines = True
```


## Methods

- [CategoryCollection](Excel.chartgroup.categorycollection.md)
- [FullCategoryCollection](Excel.chartgroup.fullcategorycollection.md)
- [SeriesCollection](Excel.ChartGroup.SeriesCollection.md)

## Properties

- [Application](Excel.ChartGroup.Application.md)
- [AxisGroup](Excel.ChartGroup.AxisGroup.md)
- [BinsCountValue](Excel.chartgroup.binscountvalue.md)
- [BinsOverflowEnabled](Excel.chartgroup.binsoverflowenabled.md)
- [BinsOverflowValue](Excel.chartgroup.binsoverflowvalue.md)
- [BinsType](Excel.chartgroup.binstype.md)
- [BinsUnderflowEnabled](Excel.chartgroup.binsunderflowenabled.md)
- [BinsUnderflowValue](Excel.chartgroup.binsunderflowvalue.md)
- [BinWidthValue](Excel.chartgroup.binwidthvalue.md)
- [BubbleScale](Excel.ChartGroup.BubbleScale.md)
- [Creator](Excel.ChartGroup.Creator.md)
- [DoughnutHoleSize](Excel.ChartGroup.DoughnutHoleSize.md)
- [DownBars](Excel.ChartGroup.DownBars.md)
- [DropLines](Excel.ChartGroup.DropLines.md)
- [FirstSliceAngle](Excel.ChartGroup.FirstSliceAngle.md)
- [GapWidth](Excel.ChartGroup.GapWidth.md)
- [Has3DShading](Excel.ChartGroup.Has3DShading.md)
- [HasDropLines](Excel.ChartGroup.HasDropLines.md)
- [HasHiLoLines](Excel.ChartGroup.HasHiLoLines.md)
- [HasRadarAxisLabels](Excel.ChartGroup.HasRadarAxisLabels.md)
- [HasSeriesLines](Excel.ChartGroup.HasSeriesLines.md)
- [HasUpDownBars](Excel.ChartGroup.HasUpDownBars.md)
- [HiLoLines](Excel.ChartGroup.HiLoLines.md)
- [Index](Excel.ChartGroup.Index.md)
- [Overlap](Excel.ChartGroup.Overlap.md)
- [Parent](Excel.ChartGroup.Parent.md)
- [RadarAxisLabels](Excel.ChartGroup.RadarAxisLabels.md)
- [SecondPlotSize](Excel.ChartGroup.SecondPlotSize.md)
- [SeriesLines](Excel.ChartGroup.SeriesLines.md)
- [ShowNegativeBubbles](Excel.ChartGroup.ShowNegativeBubbles.md)
- [SizeRepresents](Excel.ChartGroup.SizeRepresents.md)
- [SplitType](Excel.ChartGroup.SplitType.md)
- [SplitValue](Excel.ChartGroup.SplitValue.md)
- [UpBars](Excel.ChartGroup.UpBars.md)
- [VaryByCategories](Excel.ChartGroup.VaryByCategories.md)


## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]