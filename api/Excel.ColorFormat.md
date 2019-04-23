---
title: ColorFormat object (Excel)
keywords: vbaxl10.chm105000
f1_keywords:
- vbaxl10.chm105000
ms.prod: excel
api_name:
- Excel.ColorFormat
ms.assetid: 9bb6bc1f-9886-d290-a336-068f84cad1a9
ms.date: 03/29/2019
localization_priority: Normal
---


# ColorFormat object (Excel)

Represents the color of a one-color object, the foreground or background color of an object with a gradient or patterned fill, or the pointer color.


## Remarks

You can set colors to an explicit red-green-blue value (by using the **RGB** property) or to a color in the color scheme (by using the **SchemeColor** property).

Use one of the properties listed in the following table to return a **ColorFormat** object.

|Use this property|With this object|To return a ColorFormat object that represents this color|
|:-----|:-----|:-----|
|**[BackColor](excel.fillformat.backcolor.md)**|**[FillFormat](excel.fillformat.md)**|The background fill color (used in a shaded or patterned fill)|
|**[ForeColor](excel.fillformat.forecolor.md)**|**FillFormat**|The foreground fill color (or simply the fill color for a solid fill)|
|**[BackColor](excel.lineformat.backcolor.md)**|**[LineFormat](excel.lineformat.md)**|The background line color (used in a patterned line)|
|**[ForeColor](excel.lineformat.forecolor.md)**|**LineFormat**|The foreground line color (or just the line color for a solid line)|
|**[ForeColor](excel.shadowformat.forecolor.md)**|**[ShadowFormat](excel.shadowformat.md)**|The shadow color|
|**[ExtrusionColor](excel.threedformat.extrusioncolor.md)**|**[ThreeDFormat](excel.threedformat.md)**|The color of the sides of an extruded object|

## Example

Use the **RGB** property to set a color to an explicit red-green-blue value. The following example adds a rectangle to _myDocument_ and then sets the foreground color, background color, and gradient for the rectangle's fill.

```vb
Set myDocument = Worksheets(1) 
With myDocument.Shapes.AddShape(msoShapeRectangle, _ 
 90, 90, 90, 50).Fill 
 .ForeColor.RGB = RGB(128, 0, 0) 
 .BackColor.RGB = RGB(170, 170, 170) 
 .TwoColorGradient msoGradientHorizontal, 1 
End With
```


## Properties

- [Application](Excel.ColorFormat.Application.md)
- [Brightness](Excel.ColorFormat.Brightness.md)
- [Creator](Excel.ColorFormat.Creator.md)
- [ObjectThemeColor](Excel.ColorFormat.ObjectThemeColor.md)
- [Parent](Excel.ColorFormat.Parent.md)
- [RGB](Excel.ColorFormat.RGB.md)
- [SchemeColor](Excel.ColorFormat.SchemeColor.md)
- [TintAndShade](Excel.ColorFormat.TintAndShade.md)
- [Type](Excel.ColorFormat.Type.md)

## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]