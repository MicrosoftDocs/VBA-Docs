---
title: ShapeNodes.SetPosition method (Publisher)
keywords: vbapb10.chm3473428
f1_keywords:
- vbapb10.chm3473428
ms.prod: publisher
api_name:
- Publisher.ShapeNodes.SetPosition
ms.assetid: f1a3bf8c-9778-b994-9c79-55987c6fa632
ms.date: 06/14/2019
localization_priority: Normal
---


# ShapeNodes.SetPosition method (Publisher)

Sets the position of the specified node. Depending on the editing type of the node, this method may affect the position of adjacent nodes.


## Syntax

_expression_.**SetPosition** (_Index_, _X1_, _Y1_)

_expression_ A variable that represents a **[ShapeNodes](Publisher.ShapeNodes.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_Index_|Required| **Integer**|The node whose position is to be set. Must be a number from 1 to the number of nodes in the specified shape; otherwise, an error occurs.|
|_X1_|Required| **Variant**|The horizontal position of the node relative to the upper-left corner of the page.|
|_Y1_|Required| **Variant**|The vertical position of the node relative to the upper-left corner of the page.|

## Remarks

For the _X1_ and _Y1_ arguments, numeric values are evaluated in [points](../language/glossary/vbe-glossary.md#point); strings can be in any units supported by Microsoft Publisher (for example, "2.5 in").


## Example

This example moves the second node in the third shape in the active publication 200 points to the right and 300 points down. The shape must be a freeform drawing.

```vb
Dim arrPoints As Variant 
Dim intX As Integer 
Dim intY As Integer 
 
With ActiveDocument.Pages(1).Shapes(3).Nodes 
 arrPoints = .Item(2).Points 
 intX = arrPoints(1, 1) 
 intY = arrPoints(1, 2) 
 .SetPosition Index:=2, X1:=intX + 200, Y1:=intY + 300 
End With 

```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]