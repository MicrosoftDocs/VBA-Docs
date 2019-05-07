---
title: FillFormat.ForeColor property (Word)
keywords: vbawd10.chm164102245
f1_keywords:
- vbawd10.chm164102245
ms.prod: word
api_name:
- Word.FillFormat.ForeColor
ms.assetid: 23ee2f7e-12f4-fba4-8b8c-9d6d4debe526
ms.date: 06/08/2017
localization_priority: Normal
---


# FillFormat.ForeColor property (Word)

Returns or sets a  **[ColorFormat](Word.ColorFormat.md)** object that represents the foreground color for the fill. Read/write.


## Syntax

_expression_.**ForeColor**

_expression_ A variable that represents a **[FillFormat](word.fillformat.md)** object.


## Example

This example adds a rectangle to the active document and then sets the foreground color, background color, and gradient for the rectangle's fill.


```vb
Dim docActive As Document 
 
Set docActive = ActiveDocument
```


```vb
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