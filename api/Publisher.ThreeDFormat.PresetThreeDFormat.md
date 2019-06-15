---
title: ThreeDFormat.PresetThreeDFormat property (Publisher)
keywords: vbapb10.chm3801352
f1_keywords:
- vbapb10.chm3801352
ms.prod: publisher
api_name:
- Publisher.ThreeDFormat.PresetThreeDFormat
ms.assetid: da0b2e3e-57e5-9c6f-6d08-3f60d38ba1f8
ms.date: 06/15/2019
localization_priority: Normal
---


# ThreeDFormat.PresetThreeDFormat property (Publisher)

Returns an **[MsoPresetThreeDFormat](Office.MsoPresetThreeDFormat.md)** constant that represents the preset extrusion format. Read-only.


## Syntax

_expression_.**PresetThreeDFormat**

_expression_ A variable that represents a **[ThreeDFormat](Publisher.ThreeDFormat.md)** object.


## Return value

MsoPresetThreeDFormat


## Remarks

The **PresetThreeDFormat** property value can be one of the **MsoPresetThreeDFormat** constants declared in the Microsoft Office type library.

Each preset extrusion format contains a set of preset values for the various properties of the extrusion. If the extrusion has a custom format rather than a preset format, this property returns **msoPresetThreeDFormatMixed**. 

The values for this property correspond to the options (numbered from left to right, top to bottom) displayed when you choose the **3D Style** button on the **Formatting** toolbar.

Use the **[SetThreeDFormat](Publisher.ThreeDFormat.SetThreeDFormat.md)** method to set the preset extrusion format.


## Example

This example sets the extrusion format for the first shape on the first page of the active publication to 3D Style 12 if the shape initially has a custom extrusion format. For this example to work, the specified shape must be a 3D shape.

```vb
Sub SetPreset3D() 
 With ActiveDocument.Pages(1).Shapes(1).ThreeD 
 If .PresetThreeDFormat = msoPresetThreeDFormatMixed Then 
 .SetThreeDFormat msoThreeD12 
 End If 
 End With 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]