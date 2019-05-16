---
title: TextFrame.MarginLeft property (Word)
keywords: vbawd10.chm162660453
f1_keywords:
- vbawd10.chm162660453
ms.prod: word
api_name:
- Word.TextFrame.MarginLeft
ms.assetid: 4f4deac5-a4be-af1d-72ac-f9c3481b0290
ms.date: 06/08/2017
localization_priority: Normal
---


# TextFrame.MarginLeft property (Word)

Returns or sets the distance (in [points](../language/glossary/vbe-glossary.md#point)) between the left edge of the text frame and the left edge of the inscribed rectangle of the shape that contains the text. Read/write  **Single**.


## Syntax

_expression_.**MarginLeft**

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