---
title: ColorStop.ThemeColor property (Excel)
keywords: vbaxl10.chm851075
f1_keywords:
- vbaxl10.chm851075
ms.prod: excel
api_name:
- Excel.ColorStop.ThemeColor
ms.assetid: bb650754-204a-3618-d431-bf7db289ceeb
ms.date: 04/20/2019
localization_priority: Normal
---


# ColorStop.ThemeColor property (Excel)

Returns or sets the theme color of the represented object. Read/write.


## Syntax

_expression_.**ThemeColor**

_expression_ A variable that represents a **[ColorStop](Excel.ColorStop.md)** object.


## Return value

Long


## Example

Applies theme color to the active selection.

```vb
Range("A1:A10").Select 
With Selection.Interior.Gradient.ColorStop.Add(1) 
 .ThemeColor = xlThemeColorAccent1 
 .TintAndShade = 0 
End With
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]