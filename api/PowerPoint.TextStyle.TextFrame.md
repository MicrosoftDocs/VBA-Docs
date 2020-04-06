---
title: TextStyle.TextFrame property (PowerPoint)
keywords: vbapp10.chm579004
f1_keywords:
- vbapp10.chm579004
ms.prod: powerpoint
api_name:
- PowerPoint.TextStyle.TextFrame
ms.assetid: f7089a7e-add2-1fbc-f5c9-632bcd080597
ms.date: 06/08/2017
localization_priority: Normal
---


# TextStyle.TextFrame property (PowerPoint)

Returns a **[TextFrame](PowerPoint.TextFrame.md)** object that contains the alignment and anchoring properties for the specified shape or master text style.


## Syntax

_expression_.**TextFrame**

_expression_ A variable that represents a [TextStyle](PowerPoint.TextStyle.md) object.


## Return value

TextFrame


## Remarks

Use the  **TextRange** property of the **TextFrame** object to return the text in the text frame.

Use the  **HasTextFrame** property to determine whether a shape contains a text frame before you apply the **TextFrame** property.


## Example

This example adds a rectangle to _myDocument_, adds text to the rectangle, and sets the top margin for the text frame.


```vb
Set myDocument = ActivePresentation.Slides(1)

With myDocument.Shapes _
        .AddShape(msoShapeRectangle, 180, 175, 350, 140).TextFrame
    .TextRange.Text = "Here is some test text"
    .MarginTop = 10
End With
```


## See also


[TextStyle Object](PowerPoint.TextStyle.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]