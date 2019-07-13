---
title: ThreeDFormat.SetThreeDFormat method (PowerPoint)
keywords: vbapp10.chm557005
f1_keywords:
- vbapp10.chm557005
ms.prod: powerpoint
api_name:
- PowerPoint.ThreeDFormat.SetThreeDFormat
ms.assetid: 9685d3f9-467a-8b11-144a-c4260bdbbddd
ms.date: 06/08/2017
localization_priority: Normal
---


# ThreeDFormat.SetThreeDFormat method (PowerPoint)

Sets the preset extrusion format.


## Syntax

_expression_. `SetThreeDFormat`( `_PresetThreeDFormat_` )

_expression_ A variable that represents a [ThreeDFormat](PowerPoint.ThreeDFormat.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _PresetThreeDFormat_|Required|**MsoPresetThreeDFormat**|Specifies a preset extrusion format that corresponds to one of the options (numbered from left to right, from top to bottom) displayed when you click the  **3D Rotation** submenu on the **Shape Effects** menu.|

## Remarks

Each preset extrusion format contains a set of preset values for the various properties of the extrusion.

This method sets the  **[PresetThreeDFormat](PowerPoint.ThreeDFormat.PresetThreeDFormat.md)** property to the format specified by the PresetThreeDFormat parameter.

The value of the PresetThreeDFormat parameter can be one of these  **MsoPresetThreeDFormat** constants. Specifying **msoPresetThreeDFormatMixed** causes an error.


||
|:-----|
|**msoThreeD1**|
|**msoThreeD2**|
|**msoThreeD3**|
|**msoThreeD4**|
|**msoThreeD5**|
|**msoThreeD6**|
|**msoThreeD7**|
|**msoThreeD8**|
|**msoThreeD9**|
|**msoThreeD10**|
|**msoThreeD11**|
|**msoThreeD12**|
|**msoThreeD13**|
|**msoThreeD14**|
|**msoThreeD15**|
|**msoThreeD16**|
|**msoThreeD17**|
|**msoThreeD18**|
|**msoThreeD19**|
|**msoThreeD20**|

## Example

This example adds an oval to _myDocument_ and sets its extrusion format to 3D Style 12.


```vb
Set myDocument = ActivePresentation.Slides(1)

With myDocument.Shapes _
        .AddShape(msoShapeOval, 30, 30, 50, 25).ThreeD
    .Visible = True
    .SetThreeDFormat msoThreeD12
End With
```


## See also


[ThreeDFormat Object](PowerPoint.ThreeDFormat.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]