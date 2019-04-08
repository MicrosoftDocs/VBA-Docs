---
title: PresetTextured method (Excel Graph)
keywords: vbagr10.chm3077629
f1_keywords:
- vbagr10.chm3077629
ms.prod: excel
api_name:
- Excel.PresetTextured
ms.assetid: 4f6abf8c-c09e-6ef8-abb1-0cc643e6458b
ms.date: 04/09/2019
localization_priority: Normal
---


# PresetTextured method (Excel Graph)

Sets the format of the specified fill to a preset texture.

## Syntax

_expression_.**PresetTextured** (_PresetTexture_)

_expression_ Required. An expression that returns one of the objects in the **Applies To** list.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_PresetTexture_ |Required |**[MsoPresetTexture](office.msopresettexture.md)**|The preset texture for the specified fill. Can be one of the **MsoPresetTexture** constants.|

## Example

This example changes the chart's textured fill format from oak to walnut.

```vb
With myChart.ChartArea.Fill 
 If .Type = msoFillTextured Then 
 If .TextureType = msoTexturePreset Then 
 If .PresetTexture = msoTextureOak Then 
 .PresetTextured msoTextureWalnut 
 End If 
 End If 
 End If 
End With
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]