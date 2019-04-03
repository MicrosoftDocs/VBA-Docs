---
title: ColorFormat object (Word)
keywords: vbawd10.chm2502
f1_keywords:
- vbawd10.chm2502
ms.prod: word
api_name:
- Word.ColorFormat
ms.assetid: 5f12793f-d847-ecf2-6cf6-39387f7f0b28
ms.date: 06/08/2017
localization_priority: Normal
---


# ColorFormat object (Word)

Represents the color of a one-color object or the foreground or background color of an object with a gradient or patterned fill. You can set colors to an explicit red-green-blue value by using the  **[RGB](Word.ColorFormat.RGB.md)** property.


## Remarks

Use one of the properties listed in the following table to return a  **ColorFormat** object.



|**Use this property**|**With this object**|**To return a ColorFormat object that represents this**|
|:-----|:-----|:-----|
|**[BackColor](Word.FillFormat.BackColor.md)**|**[FillFormat](Word.FillFormat.md)**|Background fill color (used in a shaded or patterned fill)|
|**[ForeColor](Word.FillFormat.ForeColor.md)**|**[FillFormat](Word.FillFormat.md)**|Foreground fill color (or the fill color for a solid fill)|
|**[BackColor](Word.LineFormat.BackColor.md)**|**[LineFormat](Word.LineFormat.md)**|Background line color (used in a patterned line)|
|**[ForeColor](Word.LineFormat.ForeColor.md)**|**[LineFormat](Word.LineFormat.md)**|Foreground line color (or the line color for a solid line)|
|**[ForeColor](Word.ShadowFormat.ForeColor.md)**|**[ShadowFormat](Word.ShadowFormat.md)**|Shadow color|
|**[ExtrusionColor](Word.ThreeDFormat.ExtrusionColor.md)**|**[ThreeDFormat](Word.ThreeDFormat.md)**|Color of the sides of an extruded object|

Use the  **RGB** property to set a color to an explicit red-green-blue value. The following example adds a rectangle to the active document and then sets the foreground color, background color, and gradient for the rectangle's fill.




```vb
With ActiveDocument.Shapes _ 
 .AddShape(msoShapeRectangle, 90, 90, 90, 50).Fill 
 .ForeColor.RGB = RGB(128, 0, 0) 
 .BackColor.RGB = RGB(170, 170, 170) 
 .TwoColorGradient msoGradientHorizontal, 1 
End With
```


## Properties



|Name|
|:-----|
|[Application](Word.ColorFormat.Application.md)|
|[Brightness](Word.ColorFormat.Brightness.md)|
|[Creator](Word.ColorFormat.Creator.md)|
|[ObjectThemeColor](Word.ColorFormat.ObjectThemeColor.md)|
|[Parent](Word.ColorFormat.Parent.md)|
|[RGB](Word.ColorFormat.RGB.md)|
|[TintAndShade](Word.ColorFormat.TintAndShade.md)|
|[Type](Word.ColorFormat.Type.md)|

## See also


[Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]