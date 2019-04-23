---
title: ChartGroups object (Excel)
keywords: vbaxl10.chm569072
f1_keywords:
- vbaxl10.chm569072
ms.prod: excel
api_name:
- Excel.ChartGroups
ms.assetid: 991147bc-bbb5-9f7d-a7c9-55854aa50325
ms.date: 03/29/2019
localization_priority: Normal
---


# ChartGroups object (Excel)

Represents one or more series plotted in a chart with the same format.


## Remarks

A **ChartGroups** collection is a collection of all the **[ChartGroup](Excel.ChartGroup(object).md)** objects in the specified chart. A chart contains one or more chart groups, each chart group contains one or more **[Series](Excel.Series(object).md)** objects, and each series contains one or more **[Points](Excel.Point(object).md)** objects. For example, a single chart might contain both a line chart group, containing all the series plotted with the line chart format, and a bar chart group, containing all the series plotted with the bar chart format.

Use the **[ChartGroups](excel.chart.chartgroups.md)** method of the **Chart** object to return the **ChartGroups** collection. 

The following example displays the number of chart groups on embedded chart 1 on worksheet 1.

```vb
MsgBox Worksheets(1).ChartObjects(1).Chart.ChartGroups.Count
```

Use **ChartGroups** (_index_), where _index_ is the chart-group index number, to return a single **ChartGroup** object. The following example adds drop lines to chart group 1 on chart sheet 1.

```vb
Charts(1).ChartGroups(1).HasDropLines = True
```

If the chart has been activated, you can use **ActiveChart**.

```vb
Charts(1).Activate 
ActiveChart.ChartGroups(1).HasDropLines = True
```

Because the index number for a particular chart group can change if the chart format used for that group is changed, it may be easier to use one of the named chart group shortcut methods to return a particular chart group. The **[PieGroups](excel.piegroups.md)** method returns the collection of pie chart groups in a chart, the **[LineGroups](excel.linegroups.md)** method returns the collection of line chart groups, and so on. Each of these methods can be used with an index number to return a single **ChartGroup** object, or without an index number to return a **ChartGroups** collection.

## Methods

- [Item](Excel.ChartGroups.Item.md)

## Properties

- [Application](Excel.ChartGroups.Application.md)
- [Count](Excel.ChartGroups.Count.md)
- [Creator](Excel.ChartGroups.Creator.md)
- [Parent](Excel.ChartGroups.Parent.md)


## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]