---
title: ChartSeries.ParetoLineColor property (Access)
keywords: vbaac10.chm14875
f1_keywords:
- vbaac10.chm14875
api_name:
- Access.ChartSeries.ParetoLineColor
ms.date: 01/23/2025
ms.localizationpriority: medium
---


# ChartSeries.ParetoLineColor property (Access)

Returns or sets the line color for a Pareto Modern Chart (only available in Current Channel). Read/write **String**.

Use a **[system color constant](../language/reference/user-interface-help/system-color-constants.md)** or the RGB function as shown in the example.


## Syntax

_expression_.**ParetoLineColor**

_expression_ A variable that represents a **[ChartSeries](Access.ChartSeries.md)** object.


## Example

The following example sets the Pareto line color of the first series in a collection.

```vb
With myChart.ChartSeriesCollection.Item(0)
 .ParetoLineColor = RGB(210, 250, 210)
End With
```

## See also

- [Chart object](Access.Chart.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]