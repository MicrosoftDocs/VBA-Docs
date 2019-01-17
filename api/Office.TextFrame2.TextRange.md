---
title: TextFrame2.TextRange property (Office)
ms.prod: office
api_name:
- Office.TextFrame2.TextRange
ms.assetid: 6ea3de69-5c3d-2f54-c8c6-df80dab8fa62
ms.date: 06/08/2017
localization_priority: Normal
---


# TextFrame2.TextRange property (Office)

Sets the text for a range of nodes in a SmartArt object. Read-only


## Syntax

_expression_. `TextRange`

 _expression_ An expression that returns a [TextFrame2](Office.TextFrame2.md) object.


## Example

The following example sets the text inside the first node.


```vb
smartart.AllNodes(1).TextFrame2.TextRange.Text="Node 1"
```


## See also


[TextFrame2 Object](Office.TextFrame2.md)



[TextFrame2 Object Members](./overview/Library-Reference/textframe2-members-office.md)

