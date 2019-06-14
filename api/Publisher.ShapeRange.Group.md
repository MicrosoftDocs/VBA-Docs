---
title: ShapeRange.Group method (Publisher)
keywords: vbapb10.chm2294018
f1_keywords:
- vbapb10.chm2294018
ms.prod: publisher
api_name:
- Publisher.ShapeRange.Group
ms.assetid: ca3e011f-72ea-904e-da3f-cac7fe24341d
ms.date: 06/14/2019
localization_priority: Normal
---


# ShapeRange.Group method (Publisher)

Groups the shapes in the specified shape range. Returns the grouped shapes as a single **[Shape](Publisher.Shape.md)** object.


## Syntax

_expression_.**Group**

_expression_ A variable that represents a **[ShapeRange](Publisher.ShapeRange.md)** object.


## Return value

Shape


## Remarks

The specified range must contain more than one shape, or an error occurs.

Because a group of shapes is treated as a single shape, grouping and ungrouping shapes changes the number of items in the **[Shapes](Publisher.Shapes.md)** collection and changes the index numbers of items that come after the affected items in the collection.


## Example

This example adds two shapes to the first page of the active publication, groups the two new shapes, sets the fill for the group, rotates the group, and sends the group to the back of the drawing layer.

```vb
With ActiveDocument.Pages(1).Shapes 
 
 ' Add two shapes to the page. 
 .AddShape(Type:=msoShapeCan, _ 
 Left:=50, Top:=10, Width:=100, Height:=200).Name = "shpOne" 
 .AddShape(Type:=msoShapeCube, _ 
 Left:=150, Top:=250, Width:=100, Height:=200).Name = "shpTwo" 
 
 ' Group the shapes and change the formatting for the whole group. 
 With .Range(Index:=Array("shpOne", "shpTwo")).Group 
 .Fill.PresetTextured PresetTexture:=msoTextureBlueTissuePaper 
 .Rotation = 45 
 .ZOrder ZOrderCmd:=msoSendToBack 
 End With 
 
End With 

```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]