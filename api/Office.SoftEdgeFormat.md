---
title: SoftEdgeFormat Object (Office)
ms.prod: office
api_name:
- Office.SoftEdgeFormat
ms.assetid: 9d9b34e1-03b5-9e56-b9ea-89c7ecce0370
ms.date: 06/08/2017
---


# SoftEdgeFormat Object (Office)

Represents the soft edges effect in Office graphics.


## Remarks

The Soft Edge effect creates a mask around the edge of an object and blends the object with the transparent edge. The result is a faded or "feathered"edge.


## Example

This example sets the soft edge formatting for the text for the second shape on the second slide in a PowerPoint presentation:


```vb
With ActivePresentation.Slides(1).Shapes(2) 
 With .Text.Font 
 .Size = 32 
 .Name = "Palatino" 
 .Softedgeformat = msosoftedge6 
 End With 
End With 

```


## Properties



|**Name**|
|:-----|
|[Application](Office.SoftEdgeFormat.Application.md)|
|[Creator](Office.SoftEdgeFormat.Creator.md)|
|[Radius](Office.SoftEdgeFormat.Radius.md)|
|[Type](Office.SoftEdgeFormat.Type.md)|

## See also





[Object Model Reference](./overview/reference-object-library-reference-for-office.md)
