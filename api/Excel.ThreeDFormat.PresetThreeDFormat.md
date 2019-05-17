---
title: ThreeDFormat.PresetThreeDFormat property (Excel)
keywords: vbaxl10.chm119013
f1_keywords:
- vbaxl10.chm119013
ms.prod: excel
api_name:
- Excel.ThreeDFormat.PresetThreeDFormat
ms.assetid: 678fa7f1-7cdc-ce05-98f7-bc6252eb3df1
ms.date: 05/17/2019
localization_priority: Normal
---


# ThreeDFormat.PresetThreeDFormat property (Excel)

Returns the preset extrusion format. Each preset extrusion format contains a set of preset values for the various properties of the extrusion. Read-only **[MsoPresetThreeDFormat](office.msopresetthreedformat.md)**.


## Syntax

_expression_.**PresetThreeDFormat**

_expression_ A variable that represents a **[ThreeDFormat](Excel.ThreeDFormat.md)** object.


## Remarks

If the extrusion has a custom format rather than a preset format, this property returns **msoPresetThreeDFormatMixed**.

This property is read-only. To set the preset extrusion format, use the **[SetThreeDFormat](Excel.ThreeDFormat.SetThreeDFormat.md)** method.


## Example

This example sets the extrusion format for shape one on _myDocument_ to 3D Style 12 if the shape initially has a custom extrusion format.

```vb
Set myDocument = Worksheets(1) 
With myDocument.Shapes(1).ThreeD 
 If .PresetThreeDFormat = msoPresetThreeDFormatMixed Then 
 .SetThreeDFormat msoThreeD12 
 End If 
End With
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]