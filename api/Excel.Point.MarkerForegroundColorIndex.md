---
title: Point.MarkerForegroundColorIndex property (Excel)
keywords: vbaxl10.chm576087
f1_keywords:
- vbaxl10.chm576087
ms.prod: excel
api_name:
- Excel.Point.MarkerForegroundColorIndex
ms.assetid: 00d5e240-0851-ea13-11a3-5972135ca5fa
ms.date: 05/09/2019
localization_priority: Normal
---


# Point.MarkerForegroundColorIndex property (Excel)

Returns or sets the marker foreground color as an index into the current color palette, or as one of the following **[XlColorIndex](Excel.XlColorIndex.md)** constants: **xlColorIndexAutomatic** or **xlColorIndexNone**. Applies only to line, scatter, and radar charts. Read/write **Long**.


## Syntax

_expression_.**MarkerForegroundColorIndex**

_expression_ A variable that represents a **[Point](Excel.Point(object).md)** object.


## Example

This example sets the marker background and foreground colors for the second point in series one on Chart1.

```vb
With Charts("Chart1").SeriesCollection(1).Points(2) 
 .MarkerBackgroundColorIndex = 4 'green 
 .MarkerForegroundColorIndex = 3 'red 
End With
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]