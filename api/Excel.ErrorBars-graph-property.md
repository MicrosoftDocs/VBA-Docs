---
title: ErrorBars property (Excel Graph)
ms.prod: excel
api_name:
- Excel.ErrorBars
ms.assetid: 28e7e234-3731-42b6-b8dc-f1945b30678e
ms.date: 04/10/2019
localization_priority: Normal
---


# ErrorBars property (Excel Graph)

Returns an **ErrorBars** object that represents the error bars for the series. Read-only.

## Syntax

_expression_.**ErrorBars**

_expression_ Required. An expression that returns an **[ErrorBars](Excel.ErrorBars-graph-object.md)** object.

## Example

This example sets the error bar color for series one. The example should be run on a 2D line chart that has error bars for series one.

```vb
With myChart.SeriesCollection(1)
    .ErrorBars.Border.ColorIndex = 8
End With
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]