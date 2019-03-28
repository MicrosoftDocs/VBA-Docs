---
title: ColorStops object (Excel)
keywords: vbaxl10.chm852072
f1_keywords:
- vbaxl10.chm852072
ms.prod: excel
api_name:
- Excel.ColorStops
ms.assetid: e138347b-f03c-2f50-bf61-f7f2182c9681
ms.date: 06/08/2017
localization_priority: Normal
---


# ColorStops object (Excel)

A collection of all the [ColorStop](Excel.ColorStop.md) objects for the specified series.


## Remarks

Each  **ColorStop** Object represents a color stop for gradient fill in a range or selection.


## Example

The following example shows the ColorStops with LinearGradients.


```vb
With Selection.Interior 
 .Pattern = xlPatternLinearGradient 
 .Gradient.Degree = 90 
 .Gradient.ColorStops.Clear 
End With 
 
 'adds stops after any have been cleared 
With Selection.Interior.Gradient.ColorStops.Add(0) 
 .ThemeColor = xlThemeColorDark1 
 .TintAndShade = 0 
End With 
 
With Selection.Interior.Gradient.ColorStops.Add(1) 
 .ThemeColor = xlThemeColorAccent1 
 .TintAndShade = 0 
End With
```

The following example shows the ColorStops with RectangularGradients.




```vb
With Selection.Interior 
 .Pattern = xlPatternRectangularGradient 
 .Gradient.RectangleLeft = 0 
 .Gradient.RectangleRight = 0 
 .Gradient.RectangleTop = 0 
 .Gradient.RectangleBottom = 0 
 .Gradient.ColorStops.Clear 
End With 
 
With Selection.Interior.Gradient.ColorStops.Add(0) 
 .Color = 192 
 .TintAndShade = 0 
End With 
 
With Selection.Interior.Gradient.ColorStops.Add(1) 
 .ThemeColor = xlThemeColorLight1 
 .TintAndShade = 0 
End With
```


## Methods

- [Add](Excel.ColorStops.Add.md)
- [Clear](Excel.ColorStops.Clear.md)
- [Item](Excel.ColorStops.Item.md)

## Properties

- [Application](Excel.ColorStops.Application.md)
- [Count](Excel.ColorStops.Count.md)
- [Creator](Excel.ColorStops.Creator.md)
- [Parent](Excel.ColorStops.Parent.md)

## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]