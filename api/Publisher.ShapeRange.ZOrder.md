---
title: ShapeRange.ZOrder method (Publisher)
keywords: vbapb10.chm2293808
f1_keywords:
- vbapb10.chm2293808
ms.prod: publisher
api_name:
- Publisher.ShapeRange.ZOrder
ms.assetid: 2043f78c-ab83-e719-c3b5-5d75edcf1593
ms.date: 06/14/2019
localization_priority: Normal
---


# ShapeRange.ZOrder method (Publisher)

Moves the specified shape in front of or behind other shapes in the collection (that is, changes the shape's position in the z-order).


## Syntax

_expression_.**ZOrder** (_ZOrderCmd_)

_expression_ A variable that represents a **[ShapeRange](Publisher.ShapeRange.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_ZOrderCmd_|Required| **[MsoZOrderCmd](office.msozordercmd.md)**|Specifies where to move the specified shape relative to the other shapes. Can be one of the **MsoZOrderCmd** constants declared in the Microsoft Office type library.|

## Return value

Nothing


## Remarks

To determine a shape's current position in the z-order, use the **[ZOrderPosition](Publisher.ShapeRange.ZOrderPosition.md)** property. 


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

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]