---
title: PresetGradient method (Excel Graph)
keywords: vbagr10.chm67172
f1_keywords:
- vbagr10.chm67172
ms.prod: excel
api_name:
- Excel.PresetGradient
ms.assetid: db282722-c2ad-b504-62b3-326814fd8ca0
ms.date: 04/09/2019
localization_priority: Normal
---


# PresetGradient method (Excel Graph)

Sets the specified fill to a preset gradient.

## Syntax

_expression_.**PresetGradient** (_Style_, _Variant_, _PresetGradientType_)

_expression_ Required. An expression that returns one of the objects in the **Applies To** list.

## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_Style_ |Required |**[MsoGradientStyle](office.msogradientstyle.md)** |The gradient style for the specified fill. Can be one of the **MsoGradientStyle** constants.|
|_Variant_ |Required |**Long**| The gradient variant for the specified fill. Can be a value from 1 through 4, corresponding to the four variants listed on the **Gradient** tab in the **Fill Effects** dialog box. If _Style_ is **msoGradientFromCenter**, the _Variant_ argument can only be 1 or 2.|
|_PresetGradientType_ |Required |**[MsoPresetGradientType](office.msopresetgradienttype.md)**|The gradient type for the specified fill. Can be one of the **MsoPresetGradientType** constants.|

## Example

This example sets the chart's fill format to the preset brass color.

```vb
With myChart.ChartArea.Fill 
 .Visible = True 
 .PresetGradient msoGradientDiagonalDown, 3, msoGradientBrass 
End With
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]