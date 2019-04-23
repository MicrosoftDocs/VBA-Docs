---
title: MarkerForegroundColorIndex property (Excel Graph)
keywords: vbagr10.chm65612
f1_keywords:
- vbagr10.chm65612
ms.prod: excel
api_name:
- Excel.MarkerForegroundColorIndex
ms.assetid: 82f8a746-821d-1349-be7a-89211387a97e
ms.date: 04/11/2019
localization_priority: Normal
---


# MarkerForegroundColorIndex property (Excel Graph)

Returns or sets the marker foreground color as an index into the current color palette, or as one of the **[XlColorIndex](excel.xlcolorindex.md)** constants. Read/write **XlColorIndex**.

## Syntax

_expression_.**MarkerForegroundColorIndex**

_expression_ Required. An expression that returns one of the objects in the **Applies To** list.

## Example

This example sets the marker background and foreground colors for the second point in series one.

```vb
With myChart.SeriesCollection(1).Points(2) 
 .MarkerBackgroundColorIndex = 4 'green 
 .MarkerForegroundColorIndex = 3 'red 
End With
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]