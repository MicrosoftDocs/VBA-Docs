---
title: ChartSeries.ComboChartType property (Access)
keywords: vbaac10.chm14781
f1_keywords:
- vbaac10.chm14781
ms.prod: access
api_name:
- Access.ChartSeries.ComboChartType
ms.date: 05/02/2018
---


# ChartSeries.ComboChartType property (Access)

Returns or sets the chart type for the specified series. Read/write **[AcChartType](Access.AcChartType.md)**.

This setting is only applicable when the **[ChartType](Access.Chart.ChartType.md)** of the parent **[Chart](Access.Chart.md)** object is set to **acChartCombo**.


## Syntax

 _expression_ . **ComboChartType**

 _expression_ A variable that represents a **ChartSeries** object.


## Example

This example checks if a chart is a combo chart, and if so, sets the **ComboChartType** of the first series to **acChartLine**.

```vb
With myChart
 If .ChartType = acChartCombo Then
  .ChartSeriesCollection.Item(0).ComboChartType = acChartLine
 End If
End With
```

## See also


#### Concepts


[AcChartType Enumeration](Access.AcChartType.md)

[ChartSeries object](Access.ChartSeries.md)

[Chart object](Access.Chart.md)