---
title: SmartArtNode.TextFrame2 property (Office)
ms.prod: office
api_name:
- Office.SmartArtNode.TextFrame2
ms.assetid: 550a5bd1-bb9d-3ffb-ed14-4687dfcc3f62
ms.date: 06/08/2017
---


# SmartArtNode.TextFrame2 property (Office)

Returns the text associated with the **SmartArtNode** object. Read-only.


## Syntax

_expression_. `TextFrame2`

_expression_ An expression that returns a [SmartArtNode](Office.SmartArtNode.md) object.


## Example

The following example sets the text inside the first node.


```vb
smartart.AllNodes(1).TextFrame2.TextRange.Text="Node 1"
```


## See also

- [SmartArtNode Object](Office.SmartArtNode.md)
- [SmartArtNode Object Members](./overview/Library-Reference/smartartnode-members-office.md)

