---
title: SmartArt.Nodes Property (Office)
ms.prod: office
api_name:
- Office.SmartArt.Nodes
ms.assetid: 0495f433-9239-a3fc-e7e9-ec79bbcc75ec
ms.date: 06/08/2017
---


# SmartArt.Nodes Property (Office)

Retrieves the children of the root node of the SmartArt diagram. Read-only


## Syntax

 _expression_. `Nodes`

 _expression_ An expression that returns a [SmartArt](./Office.SmartArt.md) object.


## Remarks

The root node has no parent node and only contains children if there are children present in the SmartArt graphic's data model. In the following example, the nodes A and F will be returned.


- A
    
- 
      - B
    
  - 
      - C
    
  - D
    
- 
      - 
      - E
    
- F
    

## Example

The following code adds a top level node in Microsoft PowerPoint.


```vb
ActivePresentation.Slides(1).Shapes(1).SmartArt.Nodes.Add
```


## See also


[SmartArt Object](Office.SmartArt.md)



[SmartArt Object Members](./overview/smartart-members-office.md)

