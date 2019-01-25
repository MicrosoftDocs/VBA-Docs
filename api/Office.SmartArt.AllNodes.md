---
title: SmartArt.AllNodes property (Office)
ms.prod: office
api_name:
- Office.SmartArt.AllNodes
ms.assetid: 8562a464-61dd-e019-9f44-89ade4703589
ms.date: 01/25/2019
localization_priority: Normal
---


# SmartArt.AllNodes property (Office)

Retrieves a **[SmartArtNodes](office.smartartnodes.md)** object containing all of the nodes within the SmartArt diagram. Read-only.


## Syntax

_expression_.**AllNodes**

_expression_ An expression that returns a **[SmartArt](Office.SmartArt.md)** object.


## Remarks

The nodes are retrieved in order, independent of data model. For example, the following data model would retrieve the nodes in order A, B, C, D, E, F.


- A
    
  - B
    
  - C
    
    - D
    
    - E
    
- F
    

## Example

The following example sets the text inside the first node.


```vb
smartart.AllNodes(1).TextFrame2.TextRange.Text="Node 1"
```


## See also

- [SmartArt object members](overview/Library-Reference/smartart-members-office.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]