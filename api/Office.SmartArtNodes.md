---
title: SmartArtNodes object (Office)
ms.prod: office
api_name:
- Office.SmartArtNodes
ms.assetid: 4c35e5a4-15a1-dd6d-85a2-eb30cbaa3093
ms.date: 01/25/2019
localization_priority: Normal
---


# SmartArtNodes object (Office)

Represents a collection of nodes within a SmartArt diagram. 


## Remarks

These nodes correspond directly to semantic elements contained within the data model of the graphic.


## Example

The following code returns the number of nodes in the SmartArt diagram.


```vb
ActivePresentation.Slides(1).Shapes(1).SmartArtNodes.Count
```
## See also

- [SmartArtNodes object members](overview/Library-Reference/smartartnodes-members-office.md)
- [Object Model Reference](overview/Library-Reference/reference-object-library-reference-for-office.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]