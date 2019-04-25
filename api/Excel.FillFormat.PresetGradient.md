---
title: FillFormat.PresetGradient method (Excel)
keywords: vbaxl10.chm115005
f1_keywords:
- vbaxl10.chm115005
ms.prod: excel
api_name:
- Excel.FillFormat.PresetGradient
ms.assetid: 0bcebb14-7f39-d20c-6701-76355c212f99
ms.date: 04/26/2019
localization_priority: Normal
---


# FillFormat.PresetGradient method (Excel)

Sets the specified fill to a preset gradient.


## Syntax

_expression_.**PresetGradient** (_Style_, _Variant_, _PresetGradientType_)

_expression_ A variable that represents a **[FillFormat](Excel.FillFormat.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Style_|Required| **[MsoGradientStyle](Office.MsoGradientStyle.md)**|The gradient style.|
| _Variant_|Required| **Integer**|The gradient variant. Can be a value from 1 through 4, corresponding to one of the four variants on the **Gradient** tab in the **Fill Effects** dialog box. If _Style_ is **msoGradientFromCenter**, the _Variant_ argument can only be 1 or 2.|
| _PresetGradientType_|Required| **[MsoPresetGradientType](office.msopresetgradienttype.md)**|The preset gradient type.|



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]