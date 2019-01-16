---
title: FormatColor object (Excel)
keywords: vbaxl10.chm801072
f1_keywords:
- vbaxl10.chm801072
ms.prod: excel
api_name:
- Excel.FormatColor
ms.assetid: b7818b27-8790-ef52-c24e-8edbdcf979f2
ms.date: 06/08/2017
localization_priority: Normal
---


# FormatColor object (Excel)

Represents the fill color specified for a threshold of a color scale conditional format or the color of the bar in a data bar conditional format.


## Remarks

You can choose a color by passing an RGB value in the  **[Color](Excel.FormatColor.Color.md)** property or designate the color by indexing into the theme color palette using the **[ThemeColor](Excel.FormatColor.ThemeColor.md)** property.


## Example

The following code example creates a range of numbers and then applies a two-color scale conditional formatting rule to that range. The color for the minimum threshold is then assigned to red and the maximum threshold to blue by indexing into the  **[ColorScaleCriteria](Excel.ColorScaleCriteria.md)** collection to set individual criteria.


```vb
Sub CreateColorScaleCF() 
 
 Dim cfColorScale As ColorScale 
 
 'Fill cells with sample data from 1 to 10 
 With ActiveSheet 
 .Range("C1") = 1 
 .Range("C2") = 2 
 .Range("C1:C2").AutoFill Destination:=Range("C1:C10") 
 End With 
 
 Range("C1:C10").Select 
 
 'Create a two-color ColorScale object for the created sample data range 
 Set cfColorScale = Selection.FormatConditions.AddColorScale(ColorScaleType:=2) 
 
 'Set the minimum threshold to red and maximum threshold to blue 
 cfColorScale.ColorScaleCriteria(1).FormatColor.Color = RGB(255, 0, 0) 
 cfColorScale.ColorScaleCriteria(2).FormatColor.Color = RGB(0, 0, 255) 
 
End Sub
```

## Properties

|Name|
|:-----|
|[Application Property](excel.formatcolor.application.md)|
|[Color Property](excel.formatcolor.color.md)|
|[ColorIndex Property](excel.formatcolor.colorindex.md)|
|[Creator Property](excel.formatcolor.creator.md)|
|[Parent Property](excel.formatcolor.parent.md)|
|[ThemeColor Property](excel.formatcolor.themecolor.md)|
|[TintAndShade Property](excel.formatcolor.tintandshade.md)|

## See also

[ColorScaleCriterion.FormatColor Property](excel.colorscalecriterion.formatcolor.md)

