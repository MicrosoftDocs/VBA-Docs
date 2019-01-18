---
title: TextFrame2.NoTextRotation property (Office)
ms.prod: office
api_name:
- Office.TextFrame2.NoTextRotation
ms.assetid: a20eae43-cc72-5dc5-c240-a3e9f7aa3a18
ms.date: 06/08/2017
localization_priority: Normal
---


# TextFrame2.NoTextRotation property (Office)

Returns or sets a value that specifies if the text on a shape is rotated if the shape itself is being rotated. Read/write


## Syntax

_expression_. `NoTextRotation`

 _expression_ An expression that returns a [TextFrame2](Office.TextFrame2.md) object.


## Remarks

Returns or sets MsoTriState enumerations with the following values: 


-  **msoCTrue**
    
-  **msoFalse**
    
-  **msoTriStateMixed**
    
-  **msoTriStateToggle**
    
-  **msoTrue**
    

## Example

The following example adds a rectangle to myDocument, adds text to the rectangle, sets the margins for the text frame, and then specifies that text rotation within the shape is not available.


```vb
Set myDocument = Worksheets(1) 
With myDocument.Shapes.AddShape(msoShapeRectangle, _ 
 0, 0, 250, 140).TextFrame2 
 .TextRange.Text = "Here is some test text" 
 .MarginBottom = 10 
 .MarginLeft = 10 
 .MarginRight = 10 
 .MarginTop = 10 
 .NoTextRotation = msoFalse 
End With 

```


## See also


[TextFrame2 Object](Office.TextFrame2.md)



[TextFrame2 Object Members](./overview/Library-Reference/textframe2-members-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]