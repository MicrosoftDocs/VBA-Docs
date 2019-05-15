---
title: SmartArt.Nodes property (Office)
ms.prod: office
api_name:
- Office.SmartArt.Nodes
ms.assetid: 0495f433-9239-a3fc-e7e9-ec79bbcc75ec
ms.date: 01/25/2019
localization_priority: Normal
---


# SmartArt.Nodes property (Office)

Retrieves the children of the root [node](office.smartartnode.md) of the SmartArt diagram. Read-only.

## Syntax

_expression_.**Nodes**

_expression_ An expression that returns a **[SmartArt](Office.SmartArt.md)** object.


## Remarks

The root node has no parent node and only contains children if there are children present in the SmartArt graphic's data model. In the following example, the nodes A and F will be returned.


- A
    
  - B
    
  - C
    
  - D
    
    - E
    
- F
    

## Example

The following code adds a top-level node in Microsoft PowerPoint.


```vb
ActivePresentation.Slides(1).Shapes(1).SmartArt.Nodes.Add
```


## See also

- [SmartArt object members](overview/Library-Reference/smartart-members-office.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]