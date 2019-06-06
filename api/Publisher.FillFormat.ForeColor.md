---
title: FillFormat.ForeColor property (Publisher)
keywords: vbapb10.chm2359553
f1_keywords:
- vbapb10.chm2359553
ms.prod: publisher
api_name:
- Publisher.FillFormat.ForeColor
ms.assetid: 39e7cf23-2ff8-69f3-8bf3-9051959c5418
ms.date: 06/07/2019
localization_priority: Normal
---


# FillFormat.ForeColor property (Publisher)

Returns or sets a **[ColorFormat](Publisher.ColorFormat.md)** object representing the foreground color for the fill, line, or shadow. Read/write.


## Syntax

_expression_.**ForeColor**

_expression_ A variable that represents a **[FillFormat](publisher.fillformat.md)** object.


## Remarks

Use the **[BackColor](publisher.fillformat.backcolor.md)** property to set the background color for a fill or line.


## Example

This example adds a rectangle to the active publication and then sets the foreground color, background color, and gradient for the rectangle's fill.

```vb
With ActiveDocument.Pages(1).Shapes.AddShape _ 
 (Type:=msoShapeRectangle, _ 
 Left:=90, Top:=90, Width:=90, Height:=50).Fill 
 .ForeColor.RGB = RGB(128, 0, 0) 
 .BackColor.RGB = RGB(170, 170, 170) 
 .TwoColorGradient msoGradientHorizontal, 1 
End With
```

<br/>

This example adds a patterned line to the active publication.

```vb
With ActiveDocument.Pages(1).Shapes.AddLine _ 
 (BeginX:=10, BeginY:=100, EndX:=250, EndY:=0).Line 
 .Weight = 6 
 .ForeColor.RGB = RGB(0, 0, 255) 
 .BackColor.RGB = RGB(128, 0, 0) 
 .Pattern = msoPatternDarkDownwardDiagonal 
End With 

```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]