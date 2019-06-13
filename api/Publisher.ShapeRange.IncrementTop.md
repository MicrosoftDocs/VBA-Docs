---
title: ShapeRange.IncrementTop method (Publisher)
keywords: vbapb10.chm2293794
f1_keywords:
- vbapb10.chm2293794
ms.prod: publisher
api_name:
- Publisher.ShapeRange.IncrementTop
ms.assetid: 8172406f-fac5-ad3d-49b8-cb4858d45c6d
ms.date: 06/14/2019
localization_priority: Normal
---


# ShapeRange.IncrementTop method (Publisher)

Moves the specified shape or shape range vertically by the specified distance.


## Syntax

_expression_.**IncrementTop** (_Increment_)

_expression_ A variable that represents a **[ShapeRange](Publisher.ShapeRange.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_Increment_|Required| **Variant**|The vertical distance to move the shape or shape range. A positive value moves the shape or shape range down; a negative value moves it up. Numeric values are evaluated in [points](../language/glossary/vbe-glossary.md#point); strings can be in any units supported by Microsoft Publisher (for example, "2.5 in").|

## Return value

Nothing


## Remarks

Use the **[IncrementLeft](Publisher.ShapeRange.IncrementLeft.md)** method to move shapes or shape ranges horizontally.


## Example

This example duplicates the first shape on the active publication, sets the fill for the duplicate, moves it 70 points to the right and 50 points up, and rotates it 30 degrees clockwise.

```vb
With ActiveDocument.Pages(1).Shapes(1).Duplicate 
 .Fill.PresetTextured PresetTexture:=msoTextureGranite 
 .IncrementLeft Increment:=70 
 .IncrementTop Increment:=-50 
 .IncrementRotation Increment:=30 
End With
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]