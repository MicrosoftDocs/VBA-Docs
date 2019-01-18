---
title: ThreeDFormat.PresetLightingSoftness property (Word)
keywords: vbawd10.chm164626538
f1_keywords:
- vbawd10.chm164626538
ms.prod: word
api_name:
- Word.ThreeDFormat.PresetLightingSoftness
ms.assetid: 3f33ad34-5779-63a0-fe50-a8bd0fcabe54
ms.date: 06/08/2017
localization_priority: Normal
---


# ThreeDFormat.PresetLightingSoftness property (Word)

Returns or sets the intensity of the extrusion lighting. Read/write  **MsoPresetLightingSoftness**.


## Syntax

 _expression_. `PresetLightingSoftness`

 _expression_ Required. A variable that represents a '[ThreeDFormat](Word.ThreeDFormat.md)' object.


## Example

This example specifies that the extrusion for shape one on myDocument be lit brightly from the left.


```vb
Set myDocument = ActiveDocument 
With myDocument.Shapes(1).ThreeD 
 .Visible = True 
 .PresetLightingSoftness = msoLightingBright 
 .PresetLightingDirection = msoLightingLeft 
End With
```


## See also


[ThreeDFormat Object](Word.ThreeDFormat.md)

