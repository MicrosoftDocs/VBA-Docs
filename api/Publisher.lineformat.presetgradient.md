---
title: LineFormat.PresetGradient method (Publisher)
keywords: vbapb10.chm3407889
f1_keywords:
- vbapb10.chm3407889
ms.prod: publisher
ms.assetid: 1722feb5-22d0-18dc-bae8-d6c128746f3a
ms.date: 06/08/2019
localization_priority: Normal
---


# LineFormat.PresetGradient method (Publisher)

Sets the specified line to a preset gradient.

## Syntax

_expression_.**PresetGradient** (_Style_, _Variant_, _PresetGradientType_)

_expression_ A variable that represents a **[LineFormat](Publisher.LineFormat.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_Style_|Required| **[MsoGradientStyle](Office.MsoGradientStyle.md)** |The style of the gradient. Can be one of the **MsoGradientStyle** constants declared in the Microsoft Office type library.|
|_Variant_|Required| **Long**|The gradient variant. Can be a value from 1 to 4, corresponding to the four variants on the **Gradient** tab in the **Fill Effects** dialog box. If Style is **msoGradientFromTitle** or **msoGradientFromCenter**, this argument can be either 1 or 2.|
|_PresetGradientType_|Required| **[MsoPresetGradientType](office.msopresetgradienttype.md)**|The gradient type. Can be one of the **MsoPresetGradientType** constants.|


## Return value

Void



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]