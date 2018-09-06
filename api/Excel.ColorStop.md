---
title: ColorStop Object (Excel)
keywords: vbaxl10.chm850072
f1_keywords:
- vbaxl10.chm850072
ms.prod: excel
api_name:
- Excel.ColorStop
ms.assetid: 43c4d024-8213-5f93-dfa9-229f37e09d9a
ms.date: 06/08/2017
---


# ColorStop Object (Excel)

Represents the color stop point for a gradient fill in a range or selection.


## Remarks

The  **ColorStop** object enables you to set properties for the cell fill including the **[Color](Excel.Border.Color.md)** , **[ThemeColor](Excel.Border.ThemeColor.md)** , and **[TintAndShade](Excel.Border.TintAndShade.md)** properties.


## Example

The following example shows how to apply properties to a  **ColorStop** object.

.




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


## See also


[Excel Object Model Reference](overview/Excel/object-model.md)


