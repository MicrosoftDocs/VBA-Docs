---
title: SmartArtNode.Delete method (Office)
ms.prod: office
api_name:
- Office.SmartArtNode.Delete
ms.assetid: 916b7ddb-7ec1-64d7-6c8f-0bc6de389026
ms.date: 01/25/2019
localization_priority: Normal
---


# SmartArtNode.Delete method (Office)

Removes the current SmartArt node. 


## Syntax

_expression_.**Delete**

_expression_ An expression that returns a **[SmartArtNode](Office.SmartArtNode.md)** object.


## Return value

Nothing


## Remarks

When the node is deleted, the first child gets promoted. In the following data model, if B is deleted, the data model then looks like the following: 

- A    
  - B    
    - C    
- D
    
- A
  - C
- D

    

## See also

- [SmartArtNode object members](overview/Library-Reference/smartartnode-members-office.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]