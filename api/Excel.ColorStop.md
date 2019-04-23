---
title: ColorStop object (Excel)
keywords: vbaxl10.chm850072
f1_keywords:
- vbaxl10.chm850072
ms.prod: excel
api_name:
- Excel.ColorStop
ms.assetid: 43c4d024-8213-5f93-dfa9-229f37e09d9a
ms.date: 03/29/2019
localization_priority: Normal
---


# ColorStop object (Excel)

Represents the color stop point for a gradient fill in a range or selection.


## Remarks

The **ColorStop** object enables you to set properties for the cell fill, including the **Color**, **ThemeColor**, and **TintAndShade** properties.


## Example

The following example shows how to apply properties to a **ColorStop** object.

```vb
With Selection.Interior 
 .Pattern = xlPatternLinearGradient 
 .Gradient.Degree = 135 
 .Gradient.ColorStops.Clear 
End With 
 
With Selection.Interior.Gradient.ColorStops.Add(0) 
 .ThemeColor = xlThemeColorDark1 
 .TintAndShade = 0 
End With 
 
With Selection.Interior.Gradient.ColorStops.Add(0.5) 
 .ThemeColor = xlThemeColorAccent1 
 .TintAndShade = 0 
End With 
 
With Selection.Interior.Gradient.ColorStops.Add(1) 
 .ThemeColor = xlThemeColorDark1 
 .TintAndShade = 0 
End With
```


## Methods

- [Delete](Excel.ColorStop.Delete.md)

## Properties

- [Application](Excel.ColorStop.Application.md)
- [Color](Excel.ColorStop.Color.md)
- [Creator](Excel.ColorStop.Creator.md)
- [Parent](Excel.ColorStop.Parent.md)
- [Position](Excel.ColorStop.Position.md)
- [ThemeColor](Excel.ColorStop.ThemeColor.md)
- [TintAndShade](Excel.ColorStop.TintAndShade.md)

## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]