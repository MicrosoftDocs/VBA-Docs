---
title: SmartArtNode.Nodes property (Office)
ms.prod: office
api_name:
- Office.SmartArtNode.Nodes
ms.assetid: ed1dc125-5160-ed59-3187-620e3253af59
ms.date: 01/25/2019
localization_priority: Normal
---


# SmartArtNode.Nodes property (Office)

Retrieves the children nodes associated with this SmartArt node. Read-only.


## Syntax

_expression_.**Nodes**

_expression_ An expression that returns a **[SmartArtNode](Office.SmartArtNode.md)** object.


## Example

The following code returns the number of nodes in the SmartArt diagram.

```vb
ActivePresentation.Slides(1).Shapes(1).SmartArtNodes.Nodes.Count
```


## See also

- [SmartArtNode object members](overview/Library-Reference/smartartnode-members-office.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]