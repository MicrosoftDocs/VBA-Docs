---
title: SmartArtNode.Shapes Property (Office)
ms.prod: office
api_name:
- Office.SmartArtNode.Shapes
ms.assetid: c8a6dd3f-830e-342c-39c1-a86a54c475d4
ms.date: 06/08/2017
---


# SmartArtNode.Shapes Property (Office)

Returns the shape range associated with this **SmartArtNode** object. Read-only.


## Syntax

_expression_. `Shapes`

_expression_ An expression that returns a [SmartArtNode](./Office.SmartArtNode.md) object.


## Remarks

To find the range, use the Count property.


## Example

The following code returns the shape range.


```vb
ActivePresentation.Slides(1).Shapes(1).SmartArtNodes.Item(1).Shapes.Count.
```


## See also

- [SmartArtNode Object](Office.SmartArtNode.md)
- [SmartArtNode Object Members](./overview/Library-Reference/smartartnode-members-office.md)

