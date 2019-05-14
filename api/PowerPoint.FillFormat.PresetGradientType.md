---
title: FillFormat.PresetGradientType property (PowerPoint)
keywords: vbapp10.chm552018
f1_keywords:
- vbapp10.chm552018
ms.prod: powerpoint
api_name:
- PowerPoint.FillFormat.PresetGradientType
ms.assetid: a9a4f3fc-7350-aba1-394a-10936166ea4c
ms.date: 06/08/2017
localization_priority: Normal
---


# FillFormat.PresetGradientType property (PowerPoint)

Returns the preset gradient type for the specified fill. Read-only. 


## Syntax

_expression_.**PresetGradientType**

_expression_ A variable that represents a **[FillFormat](powerpoint.fillformat.md)** object.


## Return value

MsoPresetGradientType


## Remarks

Use the  **[PresetGradient](PowerPoint.FillFormat.PresetGradient.md)** method to set the preset gradient type for the fill.

The value of the  **PresetGradientType** property can be one of these **MsoPresetGradientType** constants.


||
|:-----|
|**msoGradientBrass**|
|**msoGradientCalmWater**|
|**msoGradientChrome**|
|**msoGradientChromeII**|
|**msoGradientDaybreak**|
|**msoGradientDesert**|
|**msoGradientEarlySunset**|
|**msoGradientFire**|
|**msoGradientFog**|
|**msoGradientGold**|
|**msoGradientGoldII**|
|**msoGradientHorizon**|
|**msoGradientLateSunset**|
|**msoGradientMahogany**|
|**msoGradientMoss**|
|**msoGradientNightfall**|
|**msoGradientOcean**|
|**msoGradientParchment**|
|**msoGradientPeacock**|
|**msoGradientRainbow**|
|**msoGradientRainbowII**|
|**msoGradientSapphire**|
|**msoGradientSilver**|
|**msoGradientWheat**|
|**msoPresetGradientMixed**|

## Example

This example changes the fill for all shapes in _myDocument_ with the Moss preset gradient fill to the Fog preset gradient fill.


```vb
Set myDocument = ActivePresentation.Slides(1)

For Each s In myDocument.Shapes

    With s.Fill

        If .PresetGradientType = msoGradientMoss Then

            .PresetGradient = msoGradientFog

        End If

    End With

Next
```


## See also


[FillFormat Object](PowerPoint.FillFormat.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]