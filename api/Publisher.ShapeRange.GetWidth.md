---
title: ShapeRange.GetWidth method (Publisher)
keywords: vbapb10.chm2293785
f1_keywords:
- vbapb10.chm2293785
ms.prod: publisher
api_name:
- Publisher.ShapeRange.GetWidth
ms.assetid: a15d1b50-289a-8b02-e090-0f0a9637980a
ms.date: 06/14/2019
localization_priority: Normal
---


# ShapeRange.GetWidth method (Publisher)

Returns the width of the shape or shape range as a **Single** in the specified units. 


## Syntax

_expression_.**GetWidth** (_Unit_)

_expression_ A variable that represents a **[ShapeRange](Publisher.ShapeRange.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_Unit_|Required| **[PbUnitType](Publisher.PbUnitType.md)** |The units in which to return the width. Can be one of the **PbUnitType** constants declared in the Microsoft Publisher type library.|

## Return value

Single


## Remarks

Use the **[GetHeight](Publisher.ShapeRange.GetHeight.md)** method to return the height of a shape or shape range.


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