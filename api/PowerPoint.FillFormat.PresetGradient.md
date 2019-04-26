---
title: FillFormat.PresetGradient method (PowerPoint)
keywords: vbapp10.chm552005
f1_keywords:
- vbapp10.chm552005
ms.prod: powerpoint
api_name:
- PowerPoint.FillFormat.PresetGradient
ms.assetid: 6aa304c7-a2ee-ceea-f956-404538bebc43
ms.date: 06/08/2017
localization_priority: Normal
---


# FillFormat.PresetGradient method (PowerPoint)

Sets the specified fill to a preset gradient.


## Syntax

_expression_.**PresetGradient** (_Style_, _Variant_, _PresetGradientType_)

_expression_ A variable that represents a **[FillFormat](powerpoint.fillformat.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Style_|Required|**[MsoGradientStyle](office.msogradientstyle.md)**|The gradient style.|
| _Variant_|Required|**Integer**|The gradient variant. Can be a value from 1 to 4, corresponding to the four variants on the **Gradient** subtab on the **Shape Fill** tab. If _Style_ is **msoGradientFromTitle** or **msoGradientFromCenter**, this argument can be either 1 or 2.|
| _PresetGradientType_|Required|**[MsoPresetGradientType](office.msopresetgradienttype.md)**|The gradient type.|


## Example

This example adds a rectangle with a preset gradient fill to _myDocument_.

```vb
Set myDocument = ActivePresentation.Slides(1)

myDocument.Shapes.AddShape(msoShapeRectangle, 90, 90, 140, 80).Fill.PresetGradient msoGradientHorizontal, 1, msoGradientBrass
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]