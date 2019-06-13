---
title: ShapeRange.GetLeft method (Publisher)
keywords: vbapb10.chm2293782
f1_keywords:
- vbapb10.chm2293782
ms.prod: publisher
api_name:
- Publisher.ShapeRange.GetLeft
ms.assetid: 236717aa-368d-8403-5928-dc6c8e437c6f
ms.date: 06/14/2019
localization_priority: Normal
---


# ShapeRange.GetLeft method (Publisher)

Returns the distance of the shape's or shape range's left edge from the left edge of the leftmost page in the current view as a **Single** in the specified units.


## Syntax

_expression_.**GetLeft** (_Unit_)

_expression_ A variable that represents a **[ShapeRange](Publisher.ShapeRange.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_Unit_|Required| **[PbUnitType](Publisher.PbUnitType.md)**|The units in which to return the distance. Can be one of the **PbUnitType** constants declared in the Microsoft Publisher type library.|

## Return value

Single


## Remarks

Use the **[GetTop](Publisher.ShapeRange.GetTop.md)** method to return the distance of a shape's or shape range's top edge from the top edge of the leftmost page in the current view.


## Example

The following example displays the distances from the left and top edges of the leftmost page to the left and top edges of the shape range consisting of all the shapes on the first page. The distances are expressed in inches (to the nearest hundredth).

```vb
With ActiveDocument.Pages(1).Shapes.Range 
 MsgBox "Distance from left: " _ 
 & Format(.GetLeft(Unit:=pbUnitInch), "0.00") _ 
 & " in" & vbCr _ 
 & "Distance from top: " _ 
 & Format(.GetTop(Unit:=pbUnitInch), "0.00") _ 
 & " in" 
End With
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]