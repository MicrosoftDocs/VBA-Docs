---
title: TextFrame2.MarginLeft property (Office)
ms.prod: office
api_name:
- Office.TextFrame2.MarginLeft
ms.assetid: b50a09fd-9f81-088b-3263-d0bbb8b83379
ms.date: 06/08/2017
localization_priority: Normal
---


# TextFrame2.MarginLeft property (Office)

Returns or sets the distance (in points) between the left edge of the text frame and the left edge of the inscribed rectangle of the shape that contains the text. Read/write


## Syntax

_expression_. `MarginLeft`

 _expression_ An expression that returns a [TextFrame2](Office.TextFrame2.md) object.


## Example

The following code adds a rectangle to a slide, adds text to the rectangle, and then sets the margins for the text frame.


```vb
Set pptSlide = ActivePresentation.Slides(1) 
With pptSlide.Shapes.AddShape(msoShapeRectangle, 0, 0, 250, 140).TextFrame2 
 .TextRange.Text = "Here is some sample text" 
 .MarginBottom = 10 
 .MarginLeft = 10 
 .MarginRight = 10 
 .MarginTop = 10 
End With
```


## See also


[TextFrame2 Object](Office.TextFrame2.md)



[TextFrame2 Object Members](./overview/Library-Reference/textframe2-members-office.md)

