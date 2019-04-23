---
title: SmartArtNode.Shapes property (Office)
ms.prod: office
api_name:
- Office.SmartArtNode.Shapes
ms.assetid: c8a6dd3f-830e-342c-39c1-a86a54c475d4
ms.date: 01/25/2019
localization_priority: Normal
---


# SmartArtNode.Shapes property (Office)

Returns the shape range associated with this **SmartArtNode** object. Read-only.


## Syntax

_expression_.**Shapes**

_expression_ An expression that returns a **[SmartArtNode](Office.SmartArtNode.md)** object.


## Remarks

To find the range, use the **[Count](office.smartartnodes.count.md)** property.


## Example

The following code returns the shape range.


```vb
ActivePresentation.Slides(1).Shapes(1).SmartArtNodes.Item(1).Shapes.Count.
```


## See also

- [SmartArtNode object members](overview/Library-Reference/smartartnode-members-office.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]