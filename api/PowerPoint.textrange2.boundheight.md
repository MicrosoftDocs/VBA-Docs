---
title: TextRange2.BoundHeight property (PowerPoint)
ms.assetid: 0a31b143-3bb6-496e-ae50-acde90472742
ms.date: 06/08/2017
ms.prod: powerpoint
localization_priority: Normal
---


# TextRange2.BoundHeight property (PowerPoint)

Gets the height, in [points](../language/glossary/vbe-glossary.md#point), of the text bounding box for the specified text. Read-only.


## Syntax

_expression_. `BoundHeight`

 _expression_ An expression that returns a 'TextRange2' object.


## Return value

Single


## Remarks

The text bounding box is not the same as the  **TextFrame** object. The **TextFrame** object represents the container in which the text can reside. The text bounding box represents the perimeter immediately surrounding the text.


## Example

This example adds a rounded rectangle to slide one with the same dimensions as the text bounding box in a PowerPoint presentation.


```vb
With ActivePresentation.Slides(1).Shapes(1) 
 Set txb = .TextFrame.Text 
 Set roundRect = .AddShape(ppShapeRoundRect, _ 
 txb.BoundLeft, txb.BoundTop, txb.BoundWidth, txb.BoundHeight) 
 roundRect.Fill.Transparency = 0.25 
End With 

```


## See also


[TextRange2 object (PowerPoint)](PowerPoint.textrange2.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]