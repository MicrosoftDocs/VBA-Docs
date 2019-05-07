---
title: FillFormat.BackColor property (Word)
keywords: vbawd10.chm164102244
f1_keywords:
- vbawd10.chm164102244
ms.prod: word
api_name:
- Word.FillFormat.BackColor
ms.assetid: bea1c59d-24ed-667c-45da-90626e8ec506
ms.date: 06/08/2017
localization_priority: Normal
---


# FillFormat.BackColor property (Word)

Returns or sets a  **[ColorFormat](Word.ColorFormat.md)** object that represents the background color for the fill Read/write.


## Syntax

_expression_. `BackColor`

_expression_ A variable that represents a **[FillFormat](word.fillformat.md)** object.


## Example

This example adds a rectangle to the active document and then sets the foreground color, background color, and gradient for the rectangle's fill.


```vb
Dim docActive As Document 
 
Set docActive = ActiveDocument 
 
With docActive.Shapes.AddShape(msoShapeRectangle, _ 
 90, 90, 90, 50).Fill 
 .ForeColor.RGB = RGB(128, 0, 0) 
 .BackColor.RGB = RGB(170, 170, 170) 
 .TwoColorGradient msoGradientHorizontal, 1 
End With
```


## See also


[FillFormat Object](Word.FillFormat.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]