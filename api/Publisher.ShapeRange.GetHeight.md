---
title: ShapeRange.GetHeight method (Publisher)
keywords: vbapb10.chm2293784
f1_keywords:
- vbapb10.chm2293784
ms.prod: publisher
api_name:
- Publisher.ShapeRange.GetHeight
ms.assetid: 63501bf7-c24d-b58e-e4c5-c8a229f07c4e
ms.date: 06/14/2019
localization_priority: Normal
---


# ShapeRange.GetHeight method (Publisher)

Returns the height of the shape or shape range as a **Single** in the specified units.


## Syntax

_expression_.**GetHeight** (_Unit_)

_expression_ A variable that represents a **[ShapeRange](Publisher.ShapeRange.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_Unit_|Required| **[PbUnitType](Publisher.PbUnitType.md)** |The units in which to return the height. Can be one of the **PbUnitType** constants declared in the Microsoft Publisher type library.|

## Return value

Single


## Remarks

Use the **[GetWidth](Publisher.ShapeRange.GetWidth.md)** method to return the width of a shape or shape range.


## Example

The following example displays the height and width in inches (to the nearest hundredth) of the shape range consisting of all the shapes on the first page of the active publication.

```vb
With ActiveDocument.Pages(1).Shapes.Range 
 MsgBox "Height of all shapes: " _ 
 & Format(.GetHeight(Unit:=pbUnitInch), "0.00") _ 
 & " in" & vbCr _ 
 & "Width of all shapes: " _ 
 & Format(.GetWidth(Unit:=pbUnitInch), "0.00") _ 
 & " in" 
End With 

```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]