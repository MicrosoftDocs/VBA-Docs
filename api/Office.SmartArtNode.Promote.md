---
title: SmartArtNode.Promote method (Office)
ms.prod: office
api_name:
- Office.SmartArtNode.Promote
ms.assetid: 806dae89-7a05-7597-70dc-ad297c79fbff
ms.date: 01/25/2019
localization_priority: Normal
---


# SmartArtNode.Promote method (Office)

Promotes the current node (and all its children) a single level within the data model.


## Syntax

_expression_.**Promote**

_expression_ An expression that returns a **[SmartArtNode](Office.SmartArtNode.md)** object.


## Return value

Nothing


## Remarks

This functionality mimics the **Promote** button on the Microsoft Office Fluent Ribbon UI when working within the content pane. For example, given the following data model, if B is promoted, the resulting data model looks like the following: 

- A
  - B
    - C   
- D
    

- A    
- B    
  - C   
- D
    

## See also

- [SmartArtNode object members](overview/Library-Reference/smartartnode-members-office.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]