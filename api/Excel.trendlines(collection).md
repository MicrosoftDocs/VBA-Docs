---
title: Trendlines collection (Excel Graph)
keywords: vbagr10.chm5208077
f1_keywords:
- vbagr10.chm5208077
ms.prod: excel
ms.assetid: 4b12461a-65a2-c535-e98d-ff68ffa5919c
ms.date: 04/06/2019
localization_priority: Normal
---


# Trendlines collection (Excel Graph)

A collection of all the **[Trendline](Excel.Trendline-graph-object.md)** objects for the specified series. Each **Trendline** object represents a trendline in a chart. A trendline shows the trend, or direction, of data in a series.


## Remarks

Use the **[Trendlines](excel.trendlines-graph-method.md)** method to return the **Trendlines** collection. 

Use the **[Add](Excel.Add.md)** method to create a new trendline and add it to the series.

Use **Trendlines** (_index_), where _index_ is the trendline's index number, to return a single **TrendLine** object.

The index number denotes the order in which the trendlines are added to the series. `Trendlines(1)` is the first trendline added to the series, and `Trendlines(Trendlines.Count)` is the last one added.

## Example

The following example displays the number of trendlines for series one in the chart.

```vb
MsgBox myChart.SeriesCollection(1).Trendlines.Count
```

<br/>

The following example adds a linear trendline to series one in the chart.

```vb
With myChart.SeriesCollection(1).Trendlines 
 .Add Type:=xlLinear, Name:="Linear Trend" 
End With
```

<br/>

The following example changes the trendline type for series one in the chart. If the series has no trendline, this example fails.

```vb
myChart.SeriesCollection(1).Trendlines(1).Type = xlMovingAvg
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]