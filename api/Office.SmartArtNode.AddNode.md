---
title: SmartArtNode.AddNode method (Office)
ms.prod: office
api_name:
- Office.SmartArtNode.AddNode
ms.assetid: f3022423-4416-ab89-ff89-e6c46d65f42c
ms.date: 01/25/2019
localization_priority: Normal
---


# SmartArtNode.AddNode method (Office)

Adds a new **SmartArtNode** object to the data model in the way specified by the **SmartArtNodePosition** value, and of type **SmartArtNodeType**.


## Syntax

_expression_.**AddNode** (_Position_, _Type_)

_expression_ An expression that returns a **[SmartArtNode](Office.SmartArtNode.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Position_|Optional|**[MsoSmartArtNodePosition](office.msosmartartnodeposition.md)**|Specifies the location of the **SmartArtNode** in the data model; for example, **msoSmartArtNodeAbove** or **msoSmartArtNodeAfter**.|
| _Type_|Optional|**[MsoSmartArtNodeType](office.msosmartartnodetype.md)**|Specifies the type of the added **SmartArtNode**; for example, **msoSmartArtNodeTypeAssistant** or **msoSmartArtNodeTypeDefault**.|

## Return value

SmartArtNode


## Example

The following code adds a default **SmartArtNode** below the current node. 


```vb
Dim saNode As SmartArtNode 
 
saNode = saNode.AddNode(msoSmartArtNodeBelow, msoSmartArtNodeTypeDefault)
```


## See also

- [SmartArtNode object members](overview/Library-Reference/smartartnode-members-office.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]