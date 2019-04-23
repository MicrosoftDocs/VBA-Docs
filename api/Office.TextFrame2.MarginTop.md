---
title: TextFrame2.MarginTop property (Office)
ms.prod: office
api_name:
- Office.TextFrame2.MarginTop
ms.assetid: d42e148d-8a92-3331-b179-3a3af4447328
ms.date: 01/25/2019
localization_priority: Normal
---


# TextFrame2.MarginTop property (Office)

Returns or sets the distance (in points) between the top of the text frame and the top of the inscribed rectangle of the shape that contains the text. Read/write.


## Syntax

_expression_.**MarginTop**

_expression_ An expression that returns a **[TextFrame2](Office.TextFrame2.md)** object.


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

- [TextFrame2 object members](overview/Library-Reference/textframe2-members-office.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]