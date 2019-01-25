---
title: SmartArtNode.Level property (Office)
ms.prod: office
api_name:
- Office.SmartArtNode.Level
ms.assetid: 63143dbc-ecd2-240c-f4c1-2b32cd47872d
ms.date: 01/25/2019
localization_priority: Normal
---


# SmartArtNode.Level property (Office)

Retrieves the node's level in the hierarchy. Read-only.


## Syntax

_expression_.**Level**

_expression_ An expression that returns a **[SmartArtNode](Office.SmartArtNode.md)** object.


## Remarks

The levels start at 1 and increment upward. If a node has no level, a 0 is returned. For example, in the following data model, A and F have a level of 1, B and D have a level of 2, and C and E have a level of 3.

- A   
  - B 
    - C    
  - D    
    - E    
- F
    

## See also

- [SmartArtNode object members](overview/Library-Reference/smartartnode-members-office.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]