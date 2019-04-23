---
title: TextRange2.Parent property (PowerPoint)
ms.assetid: 0eaca5f5-de68-4d9b-96a3-0323dff39a4b
ms.date: 06/08/2017
ms.prod: powerpoint
localization_priority: Normal
---


# TextRange2.Parent property (PowerPoint)

Gets the  **Parent** object for the **TextRange2** object. Read-only.


## Syntax

_expression_.**Parent**

 _expression_ An expression that returns a 'TextRange2' object.


## Return value

Object


## Example

This example adds an oval containing text to slide one in the active presentation and rotates the oval and the text 45 degrees. The parent object for the text frame is the  **Shape** object that contains the text.


```vb
Set myShapes = ActivePresentation.Slides(1).Shapes 
With myShapes.AddShape(Type:=msoShapeOval, Left:=50, _ 
 Top:=50, Width:=300, Height:=150).TextFrame 
 .TextRange2.Text = "Test text" 
 .Parent.Rotation = 45 
End With
```


## See also


[TextRange2 object (PowerPoint)](PowerPoint.textrange2.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]