---
title: ThreeDFormat.SetThreeDFormat method (Excel)
keywords: vbaxl10.chm119003
f1_keywords:
- vbaxl10.chm119003
ms.prod: excel
api_name:
- Excel.ThreeDFormat.SetThreeDFormat
ms.assetid: 64315607-991a-426e-e931-78432558832e
ms.date: 05/17/2019
localization_priority: Normal
---


# ThreeDFormat.SetThreeDFormat method (Excel)

Sets the preset extrusion format. Each preset extrusion format contains a set of preset values for the various properties of the extrusion.


## Syntax

_expression_.**SetThreeDFormat** (_PresetThreeDFormat_)

_expression_ A variable that represents a **[ThreeDFormat](Excel.ThreeDFormat.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _PresetThreeDFormat_|Required| **[MsoPresetThreeDFormat](Office.MsoPresetThreeDFormat.md)**|Specifies a preset extrusion format that corresponds to one of the options (numbered from left to right, from top to bottom) displayed when you choose the **3D** button on the **Drawing** toolbar.|

## Remarks

This method sets the **[PresetThreeDFormat](Excel.ThreeDFormat.PresetThreeDFormat.md)** property to the format specified by the _PresetThreeDFormat_ argument.


## Example

This example adds an oval to _myDocument_ and sets its extrusion format to 3D Style 12.

```vb
Set myDocument = Worksheets(1) 
With myDocument.Shapes.AddShape(msoShapeOval, _ 
        30, 30, 50, 25).ThreeD 
    .Visible = True 
    .SetThreeDFormat msoThreeD12 
End With
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]