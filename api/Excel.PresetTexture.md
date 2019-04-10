---
title: PresetTexture property (Excel Graph)
keywords: vbagr10.chm67162
f1_keywords:
- vbagr10.chm67162
ms.prod: excel
api_name:
- Excel.PresetTexture
ms.assetid: 5b471290-66f4-3504-096b-70265db88b93
ms.date: 04/11/2019
localization_priority: Normal
---


# PresetTexture property (Excel Graph)

Returns the preset texture for the specified fill. Read-only **[MsoPresetTexture](office.msopresettexture.md)**.

## Syntax

_expression_.**PresetTexture**

_expression_ Required. An expression that returns one of the objects in the **Applies To** list.


## Remarks

This property is read-only. Use the **[PresetTextured](excel.presettextured.md)** method to set the preset texture for the fill.

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