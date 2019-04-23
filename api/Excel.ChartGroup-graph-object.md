---
title: ChartGroup object (Excel Graph)
keywords: vbagr10.chm131097
f1_keywords:
- vbagr10.chm131097
ms.prod: excel
api_name:
- Excel.ChartGroup
ms.assetid: 8a485a8c-e181-a039-60b9-a02c2c89b26e
ms.date: 04/06/2019
localization_priority: Normal
---


# ChartGroup object (Excel Graph)

Represents one or more series of points plotted in a chart with the same format. A chart contains one or more chart groups, each chart group contains one or more [series](Excel.Series-graph-object.md), and each series contains one or more [points](Excel.Point-graph-object.md). 

For example, a single chart might contain both a line chart group, which contains all the series plotted with the line chart format, and a bar chart group, which contains all the series plotted with the bar chart format. The **ChartGroup** object is a member of the **[ChartGroups](Excel.chartgroups(collection).md)** collection.


## Remarks

Use **ChartGroups** (_index_), where _index_ is the chart group's index number, to return a single **ChartGroup** object. The following example adds drop lines to chart group one in the chart.

```vb
myChart.ChartGroups(1).HasDropLines = True
```

Because the index number for a particular chart group can change if the chart format used for that group is changed, it may be easier to use one of the named shortcut methods for chart groups to return a particular chart group. The **PieGroups** method returns the collection of pie chart groups in a chart, the **LineGroups** method returns the collection of all the line chart groups, and so on. 

You can use each of these methods with an index number to return a single **ChartGroup** object, or you can use each one without an index number to return a **ChartGroups** collection. 

## Methods

- **[AreaGroups](Excel.AreaGroups.md)**     
- **[BarGroups](Excel.BarGroups.md)**    
- **[ColumnGroups](Excel.ColumnGroups.md)**  
- **[DoughnutGroups](Excel.DoughnutGroups.md)**    
- **[LineGroups](Excel.LineGroups.md)**    
- **[PieGroups](Excel.PieGroups.md)** 


## See also

- [Excel Graph Visual Basic Reference](overview/excel/graph-visual-basic-reference.md)
- [Excel Object Model Reference](overview/excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]