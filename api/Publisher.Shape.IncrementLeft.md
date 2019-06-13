---
title: Shape.IncrementLeft method (Publisher)
keywords: vbapb10.chm2228256
f1_keywords:
- vbapb10.chm2228256
ms.prod: publisher
api_name:
- Publisher.Shape.IncrementLeft
ms.assetid: 447886ad-f515-9869-524a-a803ab025fa4
ms.date: 06/13/2019
localization_priority: Normal
---


# Shape.IncrementLeft method (Publisher)

Moves the specified shape or shape range horizontally by the specified distance.


## Syntax

_expression_.**IncrementLeft** (_Increment_)

_expression_ A variable that represents a **[Shape](Publisher.Shape.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_Increment_|Required| **Variant**|The horizontal distance to move the shape or shape range. A positive value moves the shape or shape range to the right; a negative value moves it to the left. Numeric values are evaluated in [points](../language/glossary/vbe-glossary.md#point); strings can be in any units supported by Microsoft Publisher (for example, "2.5 in").|

## Return value

Nothing


## Remarks

Use the **[IncrementTop](Publisher.Shape.IncrementTop.md)** method to move shapes or shape ranges vertically.


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