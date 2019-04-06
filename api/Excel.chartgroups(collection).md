---
title: ChartGroups collection (Excel Graph)
keywords: vbagr10.chm5207191
f1_keywords:
- vbagr10.chm5207191
ms.prod: excel
ms.assetid: 203bc32b-61e7-9bbc-bcc3-c7d8afc8b2ae
ms.date: 04/06/2019
localization_priority: Normal
---


# ChartGroups collection (Excel Graph)

A collection of all the **[ChartGroup](Excel.ChartGroup-graph-object.md)** objects in the specified chart. Each **ChartGroup** object represents one or more series plotted with the same format in a chart. 

A chart contains one or more chart groups, each chart group contains one or more series, and each series contains one or more points. For example, a single chart might contain both a line chart group, containing all the series plotted with the line chart format, and a bar chart group, containing all the series plotted with the bar chart format.


## Remarks

Use the **[ChartGroups](excel.chartgroups-graph-method.md)** method to return the **ChartGroups** collection. 

Use **ChartGroups** (_index_), where _index_ is the chart group's index number, to return a single **ChartGroup** object. 

Because the index number for a particular chart group can change if the chart format used for that group is changed, it may be easier to use one of the named chart-group shortcut methods to return a particular chart group. The **PieGroups** method returns the collection of pie chart groups in the specified chart, the **LineGroups** method returns the collection of line chart groups, and so on. Each of these methods can be used with an index number to return a single **ChartGroup** object, or used without an index number to return a **ChartGroups** collection. 

## Methods

- **[AreaGroups](Excel.AreaGroups.md)**    
- **[BarGroups](Excel.BarGroups.md)**    
- **[ColumnGroups](Excel.ColumnGroups.md)**    
- **[DoughnutGroups](Excel.DoughnutGroups.md)**    
- **[LineGroups](Excel.LineGroups.md)**   
- **[PieGroups](Excel.PieGroups.md)** 

## Example

The following example displays the number of chart groups in the chart.

```vb
MsgBox myChart.ChartGroups.Count
```

<br/>

The following example adds drop lines to chart group one in the chart.

```vb
myChart.ChartGroups(1).HasDropLines = True
```


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]