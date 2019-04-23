---
title: MarkerForegroundColor property (Excel Graph)
keywords: vbagr10.chm5207661
f1_keywords:
- vbagr10.chm5207661
ms.prod: excel
api_name:
- Excel.MarkerForegroundColor
ms.assetid: 27c88341-0446-bad5-25f4-a4f19c2db4ec
ms.date: 04/11/2019
localization_priority: Normal
---


# MarkerForegroundColor property (Excel Graph)

Returns or sets the foreground color of the marker as an RGB value. Applies only to line, scatter, and radar charts. Read/write **Long**.

## Syntax

_expression_.**MarkerForegroundColor**

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