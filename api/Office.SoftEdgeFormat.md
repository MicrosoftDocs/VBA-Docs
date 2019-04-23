---
title: SoftEdgeFormat object (Office)
ms.prod: office
api_name:
- Office.SoftEdgeFormat
ms.assetid: 9d9b34e1-03b5-9e56-b9ea-89c7ecce0370
ms.date: 01/25/2019
localization_priority: Normal
---


# SoftEdgeFormat object (Office)

Represents the soft edge effect in Office graphics.


## Remarks

The soft edge effect creates a mask around the edge of an object and blends the object with the transparent edge. The result is a faded or "feathered" edge.


## Example

This example sets the soft edge formatting for the text for the second shape on the second slide in a PowerPoint presentation.


```vb
With ActivePresentation.Slides(1).Shapes(2) 
 With .Text.Font 
 .Size = 32 
 .Name = "Palatino" 
 .Softedgeformat = msosoftedge6 
 End With 
End With 

```


## See also

- [SoftEdgeFormat object members](overview/Library-Reference/softedgeformat-members-office.md)
- [Object Model Reference](overview/Library-Reference/reference-object-library-reference-for-office.md)


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]