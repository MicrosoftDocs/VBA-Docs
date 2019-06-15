---
title: ThreeDFormat.PresetLightingSoftness property (Publisher)
keywords: vbapb10.chm3801350
f1_keywords:
- vbapb10.chm3801350
ms.prod: publisher
api_name:
- Publisher.ThreeDFormat.PresetLightingSoftness
ms.assetid: 8bad53c5-9d1c-367f-3f43-64691e193334
ms.date: 06/15/2019
localization_priority: Normal
---


# ThreeDFormat.PresetLightingSoftness property (Publisher)

Returns or sets an **[MsoPresetLightingSoftness](Office.MsoPresetLightingSoftness.md)** constant that represents the intensity of the extrusion lighting. Read/write.


## Syntax

_expression_.**PresetLightingSoftness**

_expression_ A variable that represents a **[ThreeDFormat](Publisher.ThreeDFormat.md)** object.


## Return value

MsoPresetLightingSoftness


## Remarks

The **PresetLightingSoftness** property value can be one of the **MsoPresetLightingSoftness** constants declared in the Microsoft Office type library.


## Example

This example sets the extrusion for the first shape on the first page of the active publication to be lit brightly from the left. For this example to work, the specified shape must be a 3D shape.

```vb
Sub SetExtrusionLightingBrightness() 
 With ActiveDocument.Pages(1).Shapes(1).ThreeD 
 .Visible = True 
 .PresetLightingSoftness = msoLightingBright 
 .PresetLightingDirection = msoLightingLeft 
 End With 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]