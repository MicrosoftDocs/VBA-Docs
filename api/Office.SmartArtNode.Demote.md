---
title: SmartArtNode.Demote Method (Office)
ms.prod: office
api_name:
- Office.SmartArtNode.Demote
ms.assetid: 075882bd-5784-9ba3-daed-065f4bf2c86e
ms.date: 06/08/2017
---


# SmartArtNode.Demote Method (Office)

Demotes the current node a single level within the data model.


## Syntax

_expression_. `Demote`

_expression_ An expression that returns a [SmartArtNode](./Office.SmartArtNode.md) object.


### Return value

Nothing


## Remarks

This functionality mimics the Demote button in the Microsoft Office Fluent Ribbon UI when working within the content pane. For example, given the following data model, if B is demoted, the resulting data model looks like the following: 


- A    
- B   
  - C    
- D
    

- A   
  - B    
  - C   
- D
    

## See also

- [SmartArtNode Object](Office.SmartArtNode.md)
- [SmartArtNode Object Members](./overview/smartartnode-members-office.md)

