---
title: Point.MarkerBackgroundColorIndex property (Excel)
keywords: vbaxl10.chm576085
f1_keywords:
- vbaxl10.chm576085
ms.prod: excel
api_name:
- Excel.Point.MarkerBackgroundColorIndex
ms.assetid: 67201623-5c76-1983-1710-441d7e54b8a5
ms.date: 05/09/2019
localization_priority: Normal
---


# Point.MarkerBackgroundColorIndex property (Excel)

Returns or sets the marker background color as an index into the current color palette, or as one of the following **[XlColorIndex](Excel.XlColorIndex.md)** constants: **xlColorIndexAutomatic** or **xlColorIndexNone**. Applies only to line, scatter, and radar charts. Read/write **Long**.


## Syntax

_expression_.**MarkerBackgroundColorIndex**

_expression_ A variable that represents a **[Point](Excel.Point(object).md)** object.


## Example

This example sets the marker background and foreground colors for the second point in series one on Chart1.

```vb
With Charts("Chart1").SeriesCollection(1).Points(2) 
    .MarkerBackgroundColorIndex = 4    'green 
    .MarkerForegroundColorIndex = 3    'red 
End With
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]