---
title: SmartArtNode.ReorderUp method (Office)
ms.prod: office
api_name:
- Office.SmartArtNode.ReorderUp
ms.assetid: 8c33b3cc-3d28-8959-c2ec-6e38ae07fcd2
ms.date: 01/25/2019
localization_priority: Normal
---


# SmartArtNode.ReorderUp method (Office)

Swaps a node with the previous node in the bulleted list. This method reorders the node's entire family.

## Syntax

_expression_.**ReorderUp**

_expression_ An expression that returns a **[SmartArtNode](Office.SmartArtNode.md)** object.


## Return value

Nothing


## Remarks

This method simulates choosing the **Reorder Up** button on the Microsoft Office Fluent Ribbon user interface, which is located on the **SmartArt Tools** tab, on the **Design** group on **Reorder Up**.


## Example

The following code swaps the first node with the next node and reorders the node's parents.

```vb
ActivePresentation.Slides(1).Shapes(1).SmartArt.Nodes(2).ReorderUp
```


## See also

- [SmartArtNode object members](overview/Library-Reference/smartartnode-members-office.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]