---
title: ShapeRange.ZOrder Method (Publisher)
keywords: vbapb10.chm2293808
f1_keywords:
- vbapb10.chm2293808
ms.prod: publisher
api_name:
- Publisher.ShapeRange.ZOrder
ms.assetid: 2043f78c-ab83-e719-c3b5-5d75edcf1593
ms.date: 06/08/2017
---


# ShapeRange.ZOrder Method (Publisher)

Moves the specified shape in front of or behind other shapes in the collection (that is, changes the shape's position in the z-order).


## Syntax

 _expression_. **ZOrder**( **_ZOrderCmd_**)

 _expression_ A variable that represents a  **ShapeRange** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|ZOrderCmd|Required| **MsoZOrderCmd**|Specifies where to move the specified shape relative to the other shapes.|

### Return Value

Nothing


## Remarks

The ZOrderCmd parameter can be one of the  **MsoZOrderCmd** constants declared in the Microsoft Office type library and shown in the following table.



| **msoBringForward**|
| **msoBringInFrontOfText**|
| **msoBringToFront**|
| **msoSendBackward**|
| **msoSendBehindText**|
| **msoSendToBack**|

Use the  [ZOrderPosition](Publisher.Shape.ZOrderPosition.md)property to determine a shape's current position in the z-order.


## Example

This example adds an oval to the active publication and then places the oval second from the back in the z-order if there is at least one other shape on the page.


```vb
With ActiveDocument.Pages(1).Shapes _ 
 .AddShape(Type:=msoShapeOval, _ 
 Left:=100, Top:=100, Width:=100, Height:=300) 
 While .ZOrderPosition > 2 
 .ZOrder ZOrderCmd:=msoSendBackward 
 Wend 
End With 

```


