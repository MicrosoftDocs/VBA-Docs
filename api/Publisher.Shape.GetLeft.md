---
title: Shape.GetLeft method (Publisher)
keywords: vbapb10.chm2228246
f1_keywords:
- vbapb10.chm2228246
ms.prod: publisher
api_name:
- Publisher.Shape.GetLeft
ms.assetid: e8f28ab3-f9da-eae7-2a21-b8b2505e9b44
ms.date: 06/13/2019
localization_priority: Normal
---


# Shape.GetLeft method (Publisher)

Returns the distance of the shape's or shape range's left edge from the left edge of the leftmost page in the current view as a **Single** in the specified units.


## Syntax

_expression_.**GetLeft** (_Unit_)

_expression_ A variable that represents a **[Shape](Publisher.Shape.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_Unit_|Required| **[PbUnitType](Publisher.PbUnitType.md)** |The units in which to return the distance. Can be one of the **PbUnitType** constants declared in the Microsoft Publisher type library.|

## Return value

Single


## Remarks

Use the **[GetTop](Publisher.Shape.GetTop.md)** method to return the distance of a shape's or shape range's top edge from the top edge of the leftmost page in the current view.


## Example

The following example displays the distances from the left and top edges of the leftmost page to the left and top edges of a shape range consisting of all the shapes on the first page. The distances are expressed in inches (to the nearest hundredth).

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