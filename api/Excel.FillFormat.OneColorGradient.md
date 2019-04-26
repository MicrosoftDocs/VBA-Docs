---
title: FillFormat.OneColorGradient method (Excel)
keywords: vbaxl10.chm115003
f1_keywords:
- vbaxl10.chm115003
ms.prod: excel
api_name:
- Excel.FillFormat.OneColorGradient
ms.assetid: dc44ddab-7aee-acd9-1008-1a9bbae13829
ms.date: 04/26/2019
localization_priority: Normal
---


# FillFormat.OneColorGradient method (Excel)

Sets the specified fill to a one-color gradient.


## Syntax

_expression_.**OneColorGradient** (_Style_, _Variant_, _Degree_)

_expression_ A variable that represents a **[FillFormat](Excel.FillFormat.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Style_|Required| **[MsoGradientStyle](Office.MsoGradientStyle.md)**|The gradient style.|
| _Variant_|Required| **Integer**|The gradient variant. Can be a value from 1 through 4, corresponding to one of the four variants on the **Gradient** tab in the **Fill Effects** dialog box. If _GradientStyle_ is **msoGradientFromCenter**, the _Variant_ argument can only be 1 or 2.|
| _Degree_|Required| **Single**|The gradient degree. Can be a value from 0.0 (dark) through 1.0 (light).|



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]