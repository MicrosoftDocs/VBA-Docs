---
title: ChartSeries.BorderColor property (Access)
keywords: vbaac10.chm14780
f1_keywords:
- vbaac10.chm14780
ms.prod: access
api_name:
- Access.ChartSeries.BorderColor
ms.date: 11/28/2018
localization_priority: Normal
---


# ChartSeries.BorderColor property (Access)

Returns or sets the border color of a series visualization. Read/write **String**.

You can use a **[system color constant](../language/reference/user-interface-help/system-color-constants.md)** or the RGB function as shown in the example.


## Syntax

_expression_.**BorderColor**

_expression_ A variable that represents a **[ChartSeries](Access.ChartSeries.md)** object.


## Example

The following example sets the border and fill color of the first series in a collection.

```vb
With myChart.ChartSeriesCollection.Item(0)
 .BorderColor = RGB(0, 0, 0)
 .FillColor = RGB(210, 250, 210)
End With
```

## See also

- [Chart object](Access.Chart.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]