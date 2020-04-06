---
title: Shape.Nodes property (PowerPoint)
keywords: vbapp10.chm547030
f1_keywords:
- vbapp10.chm547030
ms.prod: powerpoint
api_name:
- PowerPoint.Shape.Nodes
ms.assetid: 85021d71-78f8-43e5-5a15-a0c1ae29ef61
ms.date: 06/08/2017
localization_priority: Normal
---


# Shape.Nodes property (PowerPoint)

Returns a **[ShapeNodes](PowerPoint.ShapeNodes.md)** collection that represents the geometric description of the specified shape. Applies to **Shape** objects that represent freeform drawings.


## Syntax

_expression_.**Nodes**

_expression_ A variable that represents a **[Shape](PowerPoint.Shape.md)** object.


## Example

This example adds a smooth node with a curved segment after node four in shape three on _myDocument_. Shape three must be a freeform drawing with at least four nodes.


```vb
Set myDocument = ActivePresentation.Slides(1)

With myDocument.Shapes(3).Nodes
    .Insert Index:=4, SegmentType:=msoSegmentCurve, _
        EditingType:=msoEditingSmooth, X1:=210, Y1:=100
End With
```


## See also


[Shape Object](PowerPoint.Shape.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]