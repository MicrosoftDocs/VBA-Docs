---
title: TextRange2 object (Office)
ms.prod: office
api_name:
- Office.TextRange2
ms.assetid: a6a59c9b-9b64-c1e2-2e98-a1f99025c877
ms.date: 01/25/2019
localization_priority: Normal
---


# TextRange2 object (Office)

Represents the text frame in a **Shape** or **ShapeRange** object.


## Remarks

This object contains the text in the text frame as well as the properties and methods that control the alignment and anchoring of the text frame. Use the **TextFrame2** property to return a **[TextFrame2](office.textframe2.md)** object.


## Example

The following example adds a rectangle to myDocument, adds text to the rectangle, and then sets the margins for the text frame. 


```vb
Set myDocument = Worksheets(1) 
With myDocument.Shapes.AddShape(msoShapeRectangle, _ 
 0, 0, 250, 140).TextFrame2 
 .TextRange.Text = "Here is some test text" 
 .MarginBottom = 10 
 .MarginLeft = 10 
 .MarginRight = 10 
 .MarginTop = 10 
End With 

```


## See also

- [TextRange2 object members](overview/Library-Reference/textrange2-members-office.md)
- [Object Model Reference](overview/Library-Reference/reference-object-library-reference-for-office.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]