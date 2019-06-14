---
title: ShapeNodes.SetEditingType method (Publisher)
keywords: vbapb10.chm3473427
f1_keywords:
- vbapb10.chm3473427
ms.prod: publisher
api_name:
- Publisher.ShapeNodes.SetEditingType
ms.assetid: f90b1323-d682-1b2b-6747-cea5f2cead3c
ms.date: 06/14/2019
localization_priority: Normal
---


# ShapeNodes.SetEditingType method (Publisher)

Sets the editing type of the specified node. If the node is a control point for a curved segment, this method sets the editing type of the node adjacent to it that joins two segments. Depending on the editing type, this method may affect the position of adjacent nodes.


## Syntax

_expression_.**SetEditingType** (_Index_, _EditingType_)

_expression_ A variable that represents a **[ShapeNodes](Publisher.ShapeNodes.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_Index_|Required| **Long**|The node whose editing type is to be set. Must be a number from 1 to the number of nodes in the specified shape; otherwise, an error occurs.|
|_EditingType_|Required| **[MsoEditingType](Office.MsoEditingType.md)**|The editing property of the node. Can be one of the **MsoEditingType** constants declared in the Microsoft Office type library.|


## Example

This example changes all corner nodes to smooth nodes in the third shape of the active publication. The shape must be a freeform drawing.

```vb
Dim intNode As Integer 
 
With ActiveDocument.Pages(1).Shapes(3).Nodes 
 For intNode = 1 to .Count 
 If .Item(intNode).EditingType = msoEditingCorner Then 
 .SetEditingType _ 
 Index:=intNode, EditingType:=msoEditingSmooth 
 End If 
 Next intNode 
End With 

```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]