---
title: TextFrame2 object (Office)
ms.prod: office
api_name:
- Office.TextFrame2
ms.assetid: d2903007-70d4-0b98-e617-96fb2df26975
ms.date: 01/25/2019
localization_priority: Normal
---


# TextFrame2 object (Office)

Represents the text frame in a **Shape** or **ShapeRange** object. Contains the text in the text frame and exposes properties and methods that control the alignment and anchoring of the text frame.


## Remarks

Use the **TextFrame2** property of the **Shape** and **ShapeRange** objects to return a **TextFrame2** object. 


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
- [Object Model Reference](overview/Library-Reference/reference-object-library-reference-for-office.md)


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]