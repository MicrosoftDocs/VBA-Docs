---
title: SmartArtNode.Delete Method (Office)
ms.prod: office
api_name:
- Office.SmartArtNode.Delete
ms.assetid: 916b7ddb-7ec1-64d7-6c8f-0bc6de389026
ms.date: 06/08/2017
---


# SmartArtNode.Delete Method (Office)

Removes the current SmartArt node. 


## Syntax

_expression_. `Delete`

_expression_ An expression that returns a [SmartArtNode](./Office.SmartArtNode.md) object.


### Return value

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

- [SmartArtNode Object](Office.SmartArtNode.md)
- [SmartArtNode Object Members](./overview/smartartnode-members-office.md)

