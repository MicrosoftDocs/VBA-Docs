---
title: TextFrame.MarginRight property (Word)
keywords: vbawd10.chm162660454
f1_keywords:
- vbawd10.chm162660454
ms.prod: word
api_name:
- Word.TextFrame.MarginRight
ms.assetid: 9c59758e-8813-a035-b001-5eb57371e7fd
ms.date: 06/08/2017
localization_priority: Normal
---


# TextFrame.MarginRight property (Word)

Returns or sets the distance (in [points](../language/glossary/vbe-glossary.md#point)) between the right edge of the text frame and the right edge of the inscribed rectangle of the shape that contains the text. Read/write  **Single**.


## Syntax

_expression_.**MarginRight**

 _expression_ An expression that returns a **[TextFrame](Word.TextFrame.md)** object.


## Example

This example adds a rectangle to myDocument, adds text to the rectangle, and then sets the margins for the text frame.


```vb
Set myDocument = ActiveDocument 
With myDocument.Shapes.AddShape(msoShapeRectangle, _ 
 0, 0, 250, 140).TextFrame 
 .TextRange.Text = "Here is some test text" 
 .MarginBottom = 0 
 .MarginLeft = 100 
 .MarginRight = 0 
 .MarginTop = 20 
End With
```


## See also


[TextFrame Object](Word.TextFrame.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]