---
title: MarkerBackgroundColor property (Excel Graph)
keywords: vbagr10.chm65609
f1_keywords:
- vbagr10.chm65609
ms.prod: excel
api_name:
- Excel.MarkerBackgroundColor
ms.assetid: 035d3bf9-e6cf-7f43-aaee-fc3c3926afaa
ms.date: 04/11/2019
localization_priority: Normal
---


# MarkerBackgroundColor property (Excel Graph)

Returns or sets the marker background color as an RGB value. Applies only to line, scatter, and radar charts. Read/write **Long**.

## Syntax

_expression_.**MarkerBackgroundColor**

_expression_ Required. An expression that returns one of the objects in the **Applies To** list.

## Example

This example sets the marker background and foreground colors for the second point in series one.

```vb
With myChart.SeriesCollection(1).Points(2) 
 .MarkerBackgroundColor = RGB(0,255,0) ' green 
 .MarkerForegroundColor = RGB(255,0,0) ' red 
End With
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]