---
title: SmartArtNode.ReorderDown Method (Office)
ms.prod: office
api_name:
- Office.SmartArtNode.ReorderDown
ms.assetid: 0e927b37-08b4-639d-dab3-936d1d473d20
ms.date: 06/08/2017
---


# SmartArtNode.ReorderDown Method (Office)

Swaps a node with the next node in the bulleted list. This method reorder's the nodes entire family.


## Syntax

_expression_. `ReorderDown`

_expression_ An expression that returns a [SmartArtNode](./Office.SmartArtNode.md) object.


### Return value

Nothing


## Remarks

This method simulates clicking the Reorder Down buttons on the Microsoft Office Fluent Ribbon user interface which is located on the SmartArt Tools tab, on the Design group on Reorder Down.


## Example

The following code swaps the first node with the next node and reorders all of its descendants. 

```vb
ActivePresentation.Slides(1).Shapes(1).SmartArt.Nodes(1).ReorderDown
```

## See also

- [SmartArtNode Object](Office.SmartArtNode.md)
- [SmartArtNode Object Members](./overview/smartartnode-members-office.md)

