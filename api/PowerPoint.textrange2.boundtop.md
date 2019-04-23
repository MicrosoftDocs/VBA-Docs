---
title: TextRange2.BoundTop property (PowerPoint)
ms.assetid: eefcac8c-4c48-46e5-baa4-18adf62b3abd
ms.date: 06/08/2017
ms.prod: powerpoint
localization_priority: Normal
---


# TextRange2.BoundTop property (PowerPoint)

Gets the top coordinate, in [points](../language/glossary/vbe-glossary.md#point), of the text bounding box for the specified text. Read-only.


## Syntax

_expression_. `BoundTop`

 _expression_ An expression that returns a 'TextRange2' object.


## Return value

Single


## Remarks

The text bounding box is not the same as the  **TextFrame** object. The **TextFrame** object represents the container in which the text can reside. The text bounding box represents the perimeter immediately surrounding the text.


## Example

This example adds a rounded rectangle to slide one with the same dimensions as the text bounding box.


```vb
With ActivePresentation.Slides(1).Shapes(1) 
 Set txb = .TextFrame.Text 
 Set roundRect = .AddShape(ppShapeRoundRect, _ 
 txb.BoundLeft, txb.BoundTop, txb.BoundWidth, txb.BoundHeight) 
 roundRect.Fill.Transparency = 0.25 
End With 

```


## See also


[TextRange2 Object](Office.TextRange2.md)
[TextRange2 Object Members](overview/Library-Reference/textrange2-members-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]