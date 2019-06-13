---
title: Shape.GetHeight method (Publisher)
keywords: vbapb10.chm2228248
f1_keywords:
- vbapb10.chm2228248
ms.prod: publisher
api_name:
- Publisher.Shape.GetHeight
ms.assetid: e94eaede-f2b3-4f68-b3ec-915354a1b0b7
ms.date: 06/13/2019
localization_priority: Normal
---


# Shape.GetHeight method (Publisher)

Returns the height of the shape or shape range as a **Single** in the specified units.


## Syntax

_expression_.**GetHeight** (_Unit_)

_expression_ A variable that represents a **[Shape](Publisher.Shape.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_Unit_|Required| **[PbUnitType](Publisher.PbUnitType.md)**|The units in which to return the height. Can be one of the **PbUnitType** constants declared in the Microsoft Publisher type library.|

## Return value

Single


## Remarks

Use the **[GetWidth](Publisher.Shape.GetWidth.md)** method to return the width of a shape or shape range.


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