---
title: MarkerStyle property (Excel Graph)
keywords: vbagr10.chm65608
f1_keywords:
- vbagr10.chm65608
api_name:
- Excel.MarkerStyle
ms.assetid: 6010628c-55ab-a613-efb0-53e6abb92295
ms.date: 04/11/2019
ms.localizationpriority: medium
---


# MarkerStyle property (Excel Graph)

Returns or sets the marker style for a point or series in a line chart, scatter chart, or radar chart. Read/write **[XlMarkerStyle](excel.xlmarkerstyle.md)**.

## Syntax

_expression_.**MarkerStyle**

_expression_ Required. An expression that returns one of the objects in the **Applies To** list.

## Example

This example sets the marker style for series one. The example should be run on a 2D line chart.

```vb
myChart.SeriesCollection(1).MarkerStyle = xlMarkerStyleCircle
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]