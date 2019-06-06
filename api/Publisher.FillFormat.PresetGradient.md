---
title: FillFormat.PresetGradient method (Publisher)
keywords: vbapb10.chm2359315
f1_keywords:
- vbapb10.chm2359315
ms.prod: publisher
api_name:
- Publisher.FillFormat.PresetGradient
ms.assetid: d97c4ce8-5cef-6f53-d0c8-8bcf9ab8bb80
ms.date: 06/07/2019
localization_priority: Normal
---


# FillFormat.PresetGradient method (Publisher)

Sets the specified fill to a preset gradient.


## Syntax

_expression_.**PresetGradient** (_Style_, _Variant_, _PresetGradientType_)

_expression_ A variable that represents a **[FillFormat](publisher.fillformat.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_Style_|Required| **[MsoGradientStyle](Office.MsoGradientStyle.md)** |The style of the gradient. Can be one of the **MsoGradientStyle** constants declared in the Microsoft Office type library.|
|_Variant_|Required| **Long**|The gradient variant. Can be a value from 1 to 4, corresponding to the four variants on the **Gradient** tab in the **Fill Effects** dialog box. If Style is **msoGradientFromTitle** or **msoGradientFromCenter**, this argument can be either 1 or 2.|
|_PresetGradientType_|Required| **[MsoPresetGradientType](office.msopresetgradienttype.md)**|The gradient type. Can be one of the **MsoPresetGradientType** constants.|


## Example

This example adds a rectangle with a preset gradient fill to the active publication.

```vb
ActiveDocument.Pages(1).Shapes _ 
 .AddShape(msoShapeRectangle, 90, 90, 140, 80) _ 
 .Fill.PresetGradient Style:=msoGradientHorizontal, _ 
 Variant:=1, PresetGradientType:=msoGradientBrass 

```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]