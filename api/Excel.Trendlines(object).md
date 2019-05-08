---
title: Trendlines object (Excel)
keywords: vbaxl10.chm591072
f1_keywords:
- vbaxl10.chm591072
ms.prod: excel
api_name:
- Excel.Trendlines
ms.assetid: 752cde45-c628-7550-6c88-07405821e348
ms.date: 04/02/2019
localization_priority: Normal
---


# Trendlines object (Excel)

A collection of all the **[Trendline](Excel.Trendline(object).md)** objects for the specified series.


## Remarks

Each **Trendline** object represents a trendline in a chart. A trendline shows the trend, or direction, of data in a series.


## Example

Use the **[Trendlines](Excel.Series.Trendlines.md)** method of the **Series** object to return the **Trendlines** collection. The following example displays the number of trendlines for series one on Chart1.

```vb
MsgBox Charts(1).SeriesCollection(1).Trendlines.Count
```

<br/>

Use the **Add** method to create a new trendline and add it to the series. The following example adds a linear trendline to the first series in embedded chart one on Sheet1.

```vb
Worksheets("sheet1").ChartObjects(1).Chart.SeriesCollection(1) _ 
 .Trendlines.Add type:=xlLinear, name:="Linear Trend"
```

<br/>

Use **Trendlines** (_index_), where _index_ is the trendline index number, to return a single **TrendLine** object. The following example changes the trendline type for the first series in embedded chart one on worksheet one. If the series has no trendline, this example will fail.

The index number denotes the order in which the trendlines were added to the series. `Trendlines(1)` is the first trendline added to the series, and `Trendlines(Trendlines.Count)` is the last one added.

```vb
Worksheets(1).ChartObjects(1).Chart. _ 
 SeriesCollection(1).Trendlines(1).Type = xlMovingAvg
```


## Methods

- [Add](Excel.Trendlines.Add.md)
- [Item](Excel.Trendlines.Item.md)

## Properties

- [Application](Excel.Trendlines.Application.md)
- [Count](Excel.Trendlines.Count.md)
- [Creator](Excel.Trendlines.Creator.md)
- [Parent](Excel.Trendlines.Parent.md)


## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]