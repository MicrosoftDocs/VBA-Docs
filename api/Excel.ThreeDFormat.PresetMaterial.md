---
title: ThreeDFormat.PresetMaterial property (Excel)
keywords: vbaxl10.chm119012
f1_keywords:
- vbaxl10.chm119012
ms.prod: excel
api_name:
- Excel.ThreeDFormat.PresetMaterial
ms.assetid: f9dd825a-7fb3-5d59-77d1-8ef861b9adc8
ms.date: 06/08/2017
localization_priority: Normal
---


# ThreeDFormat.PresetMaterial property (Excel)

Returns or sets the extrusion surface material. Read/write  **MsoPresetMaterial**.


## Syntax

_expression_. `PresetMaterial`

_expression_ A variable that represents a [ThreeDFormat](./Excel.ThreeDFormat.md) object.


## Remarks





| **MsoPresetMaterial** can be one of these **MsoPresetMaterial** constants.|
| **msoMaterialMatte**|
| **msoMaterialMetal**|
| **msoMaterialPlastic**|
| **msoMaterialWireFrame**|
| **msoPresetMaterialMixed**|

## Example

This example specifies that the extrusion surface for shape one in  `myDocument` be wire frame.


```vb
Set myDocument = Worksheets(1) 
With myDocument.Shapes(1).ThreeD 
 .Visible = True 
 .PresetMaterial = msoMaterialWireFrame 
End With
```


## See also


[ThreeDFormat Object](Excel.ThreeDFormat.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]