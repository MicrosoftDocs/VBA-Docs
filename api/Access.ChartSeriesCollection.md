---
title: ChartSeriesCollection object (Access)
keywords: vbaac10.chm14751
f1_keywords:
- vbaac10.chm14751
ms.prod: access
api_name:
- Access.ChartSeriesCollection
ms.date: 11/28/2018
localization_priority: Normal
---


# ChartSeriesCollection object (Access)

A collection of all the **[ChartSeries](Access.ChartSeries.md)** objects in the specified chart.


## Example

The following example displays the name of each **[ChartSeries](Access.ChartSeries.md)** instance in a collection.

```vb
With myChart
 For Each series In .ChartSeriesCollection
  MsgBox (series.Name)
 Next
End With
```

## See also

- [Chart object](Access.Chart.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]