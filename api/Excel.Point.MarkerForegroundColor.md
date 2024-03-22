---
title: Point.MarkerForegroundColor property (Excel)
keywords: vbaxl10.chm576086
f1_keywords:
- vbaxl10.chm576086
api_name:
- Excel.Point.MarkerForegroundColor
ms.assetid: 800fb100-8dc3-8e03-7308-48ffb2df552e
ms.date: 05/09/2019
ms.localizationpriority: medium
---


# Point.MarkerForegroundColor property (Excel)

Sets the marker foreground color as an RGB value or returns the corresponding color index value. The foreground color is displayed as the Border color in the application. Applies only to line, scatter, and radar charts. Read/write **Long**.


## Syntax

_expression_.**MarkerForegroundColor**

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
