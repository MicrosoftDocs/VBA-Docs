---
title: TextFrame.MarginRight property (PowerPoint)
keywords: vbapp10.chm558004
f1_keywords:
- vbapp10.chm558004
ms.prod: powerpoint
api_name:
- PowerPoint.TextFrame.MarginRight
ms.assetid: 57ab53e7-1fbf-09b6-13c4-7cb0a814d9e3
ms.date: 06/08/2017
localization_priority: Normal
---


# TextFrame.MarginRight property (PowerPoint)

Returns or sets the distance (in [points](../language/glossary/vbe-glossary.md#point)) between the right edge of the text frame and the right edge of the inscribed rectangle of the shape that contains the text. Read/write.


## Syntax

_expression_.**MarginRight**

_expression_ A variable that represents a **[TextFrame](PowerPoint.TextFrame.md)** object.


## Return value

Single


## Example

This example adds a rectangle to myDocument, adds text to the rectangle, and then sets the margins for the text frame.


```vb
Set myDocument = ActivePresentation.Slides(1)

With myDocument.Shapes.AddShape(msoShapeRectangle, _
        0, 0, 250, 140).TextFrame
    .TextRange.Text = "Here is some test text"
    .MarginBottom = 0
    .MarginLeft = 10
    .MarginRight = 5
    .MarginTop = 20
End With
```


## See also


[TextFrame Object](PowerPoint.TextFrame.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]