---
title: ThreeDFormat.PresetLightingSoftness property (Excel)
keywords: vbaxl10.chm119011
f1_keywords:
- vbaxl10.chm119011
ms.prod: excel
api_name:
- Excel.ThreeDFormat.PresetLightingSoftness
ms.assetid: e63a483b-16c6-edab-6a16-b539f0a424cb
ms.date: 05/17/2019
localization_priority: Normal
---


# ThreeDFormat.PresetLightingSoftness property (Excel)

Returns or sets the intensity of the extrusion lighting. Read/write **[MsoPresetLightingSoftness](office.msopresetlightingsoftness.md)**.


## Syntax

_expression_.**PresetLightingSoftness**

_expression_ A variable that represents a **[ThreeDFormat](Excel.ThreeDFormat.md)** object.



## Example

This example specifies that the extrusion for shape one on _myDocument_ be lit brightly from the left.

```vb
Set myDocument = Worksheets(1) 
With myDocument.Shapes(1).ThreeD 
 .Visible = True 
 .PresetLightingSoftness = msoLightingBright 
 .PresetLightingDirection = msoLightingLeft 
End With
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]