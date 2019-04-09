---
title: ThreeDFormat.PresetMaterial property (Word)
keywords: vbawd10.chm164626539
f1_keywords:
- vbawd10.chm164626539
ms.prod: word
api_name:
- Word.ThreeDFormat.PresetMaterial
ms.assetid: 95cb7421-29fb-8905-7b0e-c43ec81f6dd5
ms.date: 06/08/2017
localization_priority: Normal
---


# ThreeDFormat.PresetMaterial property (Word)

Returns or sets the extrusion surface material. Read/write  **MsoPresetMaterial**.


## Syntax

_expression_. `PresetMaterial`

_expression_ Required. A variable that represents a '[ThreeDFormat](Word.ThreeDFormat.md)' object.


## Example

This example specifies that the extrusion surface for shape one in myDocument be wireframe.


```vb
Set myDocument = ActiveDocument 
With myDocument.Shapes(1).ThreeD 
 .Visible = True 
 .PresetMaterial = msoMaterialWireFrame 
End With
```


## See also


[ThreeDFormat Object](Word.ThreeDFormat.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]