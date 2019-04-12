---
title: ColorFormat.TintAndShade property (Word)
keywords: vbawd10.chm163971175
f1_keywords:
- vbawd10.chm163971175
ms.prod: word
api_name:
- Word.ColorFormat.TintAndShade
ms.assetid: e0b54e37-475c-0e6b-f530-aa69b8fe51b8
ms.date: 06/08/2017
localization_priority: Normal
---


# ColorFormat.TintAndShade property (Word)

Returns a  **Single** that represents the lightening or darkening of a specified shape's color. Read/write.


## Syntax

_expression_.**TintAndShade**

 _expression_ An expression that returns a '[ColorFormat](Word.ColorFormat.md)' object.


## Remarks

You can enter a number from -1 (darkest) to 1 (lightest) for the TintAndShade property, 0 (zero) being neutral.


## Example

This example creates a new shape in the active document, sets the fill color, and lightens the color shade.


```vb
Sub NewTintedShape() 
 Dim shpHeart As Shape 
 Set shpHeart = ActiveDocument.Shapes _ 
 .AddShape(Type:=msoShapeHeart, Left:=150, _ 
 Top:=150, Width:=250, Height:=250) 
 With shpHeart.Fill.ForeColor 
 .RGB = RGB(Red:=255, Green:=28, Blue:=0) 
 .TintAndShade = 0.3 
 End With 
End Sub
```


## See also


[ColorFormat Object](Word.ColorFormat.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]