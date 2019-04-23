---
title: PresetGradientType property (Excel Graph)
keywords: vbagr10.chm67173
f1_keywords:
- vbagr10.chm67173
ms.prod: excel
api_name:
- Excel.PresetGradientType
ms.assetid: 10ea644f-a856-acd1-45b8-6c1d35d2390a
ms.date: 04/11/2019
localization_priority: Normal
---


# PresetGradientType property (Excel Graph)

Returns the preset gradient type for the specified fill. Read-only **[MsoPresetGradientType](office.msopresetgradienttype.md)**.

## Syntax

_expression_.**PresetGradientType**

_expression_ Required. An expression that returns one of the objects in the **Applies To** list.


## Remarks

This property is read-only. Use the **[PresetGradient](excel.presetgradient.md)** method to set the preset gradient type for the fill.

## Example

This example changes the chart's preset gradient fill format from silver to gold.

```vb
With myChart.ChartArea.Fill 
 If .Type = msoFillGradient Then 
 If .GradientColorType = msoGradientPresetColors Then 
 If .PresetGradientType = msoGradientSilver Then 
 .PresetGradient .GradientStyle, _ 
 .GradientVariant, msoGradientGold 
 End If 
 End If 
 End If 
End With
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]