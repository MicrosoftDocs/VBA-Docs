---
title: SmartArtNodes object (Office)
ms.prod: office
api_name:
- Office.SmartArtNodes
ms.assetid: 4c35e5a4-15a1-dd6d-85a2-eb30cbaa3093
ms.date: 06/08/2017
localization_priority: Normal
---


# SmartArtNodes object (Office)

Represents a collection of nodes within a Smart Art diagram. 


## Remarks

These nodes correspond directly to semantic elements contained within the data model of the graphic.


## Example

The following code returns the number of nodes in the Smart Art diagram.


```vb
ActivePresentation.Slides(1).Shapes(1).SmartArtNodes.Count
```


## Methods



|Name|
|:-----|
|[Add](Office.SmartArtNodes.Add.md)|
|[Item](Office.SmartArtNodes.Item.md)|

## Properties



|Name|
|:-----|
|[Application](Office.SmartArtNodes.Application.md)|
|[Count](Office.SmartArtNodes.Count.md)|
|[Creator](Office.SmartArtNodes.Creator.md)|
|[Parent](Office.SmartArtNodes.Parent.md)|

## See also





[Object Model Reference](./overview/Library-Reference/reference-object-library-reference-for-office.md)
