---
title: ColorFormat Object (Publisher)
keywords: vbapb10.chm2621439
f1_keywords:
- vbapb10.chm2621439
ms.prod: publisher
api_name:
- Publisher.ColorFormat
ms.assetid: 659069e1-e359-94d7-de06-a1d98378193b
ms.date: 06/08/2017
---


# ColorFormat Object (Publisher)

Represents the color of a one-color object or the foreground or background color of an object with a gradient or patterned fill. You can set colors to an explicit red-green-blue value by using the  **[RGB](Publisher.ColorFormat.RGB.md)** property.
 


## Remarks

Use one of the properties listed in the following table to return a  **ColorFormat** object.
 

 


|**Use this property**|**With this object**|**To return a ColorFormat object that represents this**|
|:-----|:-----|:-----|
|**[BackColor](fillformat-backcolor-property-publisher.md)**|**[FillFormat](fillformat-object-publisher.md)**|Background fill color (used in a shaded or patterned fill)|
|**[ForeColor](Publisher.FillFormat.ForeColor.md)**|**FillFormat**|Foreground fill color (or the fill color for a solid fill)|
|**[BackColor](lineformat-backcolor-property-publisher.md)**|**[LineFormat](lineformat-object-publisher.md)**|Background line color (used in a patterned line)|
|**[ForeColor](Publisher.LineFormat.ForeColor.md)**|**LineFormat**|Foreground line color (or the line color for a solid line)|
|**[ForeColor](shadowformat-forecolor-property-publisher.md)**|**[ShadowFormat](shadowformat-object-publisher.md)**|Shadow color|
|**[ExtrusionColor](threedformat-extrusioncolor-property-publisher.md)**|**[ThreeDFormat](threedformat-object-publisher.md)**|Color of the sides of an extruded object|

## Example

Use the  **RGB** property to set a color to an explicit red-green-blue value. The following example adds a rectangle to the active publication and then sets the foreground color, background color, and gradient for the rectangle's fill.
 

 

```
Sub GradientFill() 
 With ActiveDocument.Pages(1).Shapes _ 
 .AddShape(Type:=msoShapeRectangle, _ 
 Left:=90, Top:=90, Width:=90, Height:=50).Fill 
 .ForeColor.RGB = RGB(128, 0, 0) 
 .BackColor.RGB = RGB(170, 170, 170) 
 .TwoColorGradient msoGradientHorizontal, 1 
 End With 
End Sub
```


## Properties



|**Name**|
|:-----|
|[Application](Publisher.ColorFormat.Application.md)|
|[BaseCMYK](Publisher.ColorFormat.BaseCMYK.md)|
|[BaseRGB](Publisher.ColorFormat.BaseRGB.md)|
|[CMYK](Publisher.ColorFormat.CMYK.md)|
|[Ink](Publisher.ColorFormat.Ink.md)|
|[Parent](Publisher.ColorFormat.Parent.md)|
|[RGB](Publisher.ColorFormat.RGB.md)|
|[SchemeColor](Publisher.ColorFormat.SchemeColor.md)|
|[TintAndShade](Publisher.ColorFormat.TintAndShade.md)|
|[Transparency](Publisher.ColorFormat.Transparency.md)|
|[Type](colorformat-type-property-publisher.md)|

