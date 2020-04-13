---
title: ColorFormat object (PowerPoint)
keywords: vbapp10.chm506000
f1_keywords:
- vbapp10.chm506000
ms.prod: powerpoint
api_name:
- PowerPoint.ColorFormat
ms.assetid: 3bfcd08d-65f4-25a3-2d05-77111fbd13e5
ms.date: 06/08/2017
localization_priority: Normal
---


# ColorFormat object (PowerPoint)

Represents the color of a one-color object, the foreground or background color of an object with a gradient or patterned fill, or the pointer color. You can set colors to an explicit red-green-blue value (by using the [RGB](PowerPoint.ColorFormat.RGB.md) property) or to a color in the color scheme (by using the [SchemeColor](PowerPoint.ColorFormat.SchemeColor.md) property).


## Remarks

Use one of the properties listed in the following table to return a **ColorFormat** object.



|**Use this property**|**With this object**|**To return a ColorFormat object that represents this**|
|:-----|:-----|:-----|
|[DimColor](PowerPoint.AnimationSettings.DimColor.md)|**AnimationSettings**|Color used for dimmed objects|
|[BackColor](PowerPoint.FillFormat.BackColor.md)|**FillFormat**|Background fill color (used in a shaded or patterned fill)|
|[ForeColor](PowerPoint.FillFormat.ForeColor.md)|**FillFormat**|Foreground fill color (the fill color for a solid fill)|
|[Color](PowerPoint.Font.Color.md)|**Font**|Bullet or character color|
|[BackColor](PowerPoint.LineFormat.BackColor.md)|**LineFormat**|Background line color (used in a patterned line)|
|[ForeColor](PowerPoint.LineFormat.ForeColor.md)|**LineFormat**|Foreground line color (or just the line color for a solid line)|
|[ForeColor](PowerPoint.ShadowFormat.ForeColor.md)|**ShadowFormat**|Shadow color|
|[PointerColor](PowerPoint.SlideShowSettings.PointerColor.md)|**SlideShowSettings**|Default pointer color for a presentation|
|[PointerColor](PowerPoint.SlideShowView.PointerColor.md)|**SlideShowView**|Temporary pointer color for a view of a slide show|
|[ExtrusionColor](PowerPoint.ThreeDFormat.ExtrusionColor.md)|**ThreeDFormat**|Color of the sides of an extruded object|

## Example

Use the [SchemeColor](PowerPoint.ColorFormat.SchemeColor.md) property to set the color of a slide element to one of the colors in the standard color scheme. The following example sets the text color for shape one on slide two in the active presentation to the standard color-scheme title color.


```vb
ActivePresentation.Slides(2).Shapes(1).TextFrame.TextRange.Font.Color.SchemeColor = ppTitle
```

Use the [RGB](PowerPoint.ColorFormat.RGB.md) property to set a color to an explicit red-green-blue value. The following example adds a rectangle to _myDocument_ and then sets the foreground color, background color, and gradient for the rectangle's fill.




```vb
Set myDocument = ActivePresentation.Slides(1) 
With myDocument.Shapes.AddShape(msoShapeRectangle, 90, 90, 90, 50).Fill 
    .ForeColor.RGB = RGB(128, 0, 0) 
    .BackColor.RGB = RGB(170, 170, 170) 
    .TwoColorGradient msoGradientHorizontal, 1 
End With
```


## See also


[PowerPoint Object Model Reference](overview/PowerPoint/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]