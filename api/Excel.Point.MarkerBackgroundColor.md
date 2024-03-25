---
title: Point.MarkerBackgroundColor property (Excel)
keywords: vbaxl10.chm576084
f1_keywords:
- vbaxl10.chm576084
api_name:
- Excel.Point.MarkerBackgroundColor
ms.assetid: a283c8d2-08f2-0865-b8fe-26bc45d497d8
ms.date: 05/09/2019
ms.localizationpriority: medium
---


# Point.MarkerBackgroundColor property (Excel)

Sets the marker background color as an RGB value or returns the corresponding color index value. The background color is displayed as the Fill color in the application. Applies only to line, scatter, and radar charts. Read/write **Long**.


## Syntax

_expression_.**MarkerBackgroundColor**

_expression_ A variable that represents a **[Point](Excel.Point(object).md)** object.


## Example

This example sets the marker background (fill) and foreground (border) colors for the first point in series one on Chart1.

```vb
With Charts("Chart1").SeriesCollection(1).Points(1) 
 .MarkerBackgroundColor = RGB(0,255,0) ' green fill
 .MarkerForegroundColor = RGB(255,0,0) ' red border
End With
```

This example sets the marker colors to automatic for the second point in series one on Chart1.

```vb
With Charts("Chart1").SeriesCollection(1).Points(2) 
 .MarkerBackgroundColor = -1 ' automatic fill
 .MarkerForegroundColor = -1 ' automatic border
End With
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
