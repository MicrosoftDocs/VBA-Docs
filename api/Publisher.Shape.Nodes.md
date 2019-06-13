---
title: Shape.Nodes property (Publisher)
keywords: vbapb10.chm2228293
f1_keywords:
- vbapb10.chm2228293
ms.prod: publisher
api_name:
- Publisher.Shape.Nodes
ms.assetid: a1463ff3-5b75-e4b9-df12-985538713c7c
ms.date: 06/13/2019
localization_priority: Normal
---


# Shape.Nodes property (Publisher)

Returns a **[ShapeNodes](Publisher.ShapeNodes.md)** collection that represents the geometric description of the specified shape. Applies to **Shape** or **ShapeRange** objects that represent freeform drawings.


## Syntax

_expression_.**Nodes**

_expression_ A variable that represents a **[Shape](Publisher.Shape.md)** object.


## Example

This example adds a smooth node with a curved segment after node four in shape three on page one. Shape three must be a freeform drawing with at least four nodes.

```vb
With ActiveDocument.Pages(1) _ 
 .Shapes(3).Nodes 
 .Insert Index:=4, SegmentType:=msoSegmentCurve, _ 
 EditingType:=msoEditingSmooth, X1:=210, Y1:=100 
End With
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]