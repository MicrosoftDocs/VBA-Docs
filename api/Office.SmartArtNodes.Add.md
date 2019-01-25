---
title: SmartArtNodes.Add method (Office)
ms.prod: office
api_name:
- Office.SmartArtNodes.Add
ms.assetid: 51254d1a-0395-2b40-842c-84ba3d52a98b
ms.date: 01/25/2019
localization_priority: Normal
---


# SmartArtNodes.Add method (Office)

Adds a new **[SmartArtNode](Office.SmartArtNode.md)** object to the diagram with specified text.


## Syntax

_expression_.**Add**

_expression_ An expression that returns a **[SmartArtNodes](Office.SmartArtNodes.md)** object.


## Return value

SmartArtNode


## Remarks

This new node is added to the bottom of the data model at the top most level for this collection of nodes. If the highest level was 2, the new node's level would then be 2 as well. 


## Example

The following code adds a **SmartArtNode** to the collection of **SmartArtNodes**. 


```vb
Dim saNodes As SmartArtNodes 
saNodes.Add()
```


## See also

- [SmartArtNodes object members](overview/Library-Reference/smartartnodes-members-office.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]