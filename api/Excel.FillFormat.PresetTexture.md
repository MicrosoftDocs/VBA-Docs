---
title: FillFormat.PresetTexture Property (Excel)
keywords: vbaxl10.chm115019
f1_keywords:
- vbaxl10.chm115019
ms.prod: excel
api_name:
- Excel.FillFormat.PresetTexture
ms.assetid: 3ed8dc1b-f816-ece8-6238-44d5d8f51378
ms.date: 06/08/2017
---


# FillFormat.PresetTexture Property (Excel)

Returns the preset texture for the specified fill. Read-only  **[MsoPresetTexture](Office.MsoPresetTexture.md)** .


## Syntax

 _expression_. `PresetTexture`

 _expression_ A variable that represents a [FillFormat](Excel.FillFormat.md) object.


## Remarks

Use the  **[PresetTextured](Excel.FillFormat.PresetTextured.md)** method to set the preset texture for the fill.


## Example

This example sets the fill format for chart two to the same style used for chart one.


```vb
Set c1f = Charts(1).ChartArea.Fill 
If c1f.Type = msoFillTextured Then 
    With Charts(2).ChartArea.Fill 
        .Visible = True 
        If c1f.TextureType = msoTexturePreset Then 
            .PresetTextured c1f.PresetTexture 
        Else 
            .UserTextured c1f.TextureName 
        End If 
    End With 
End If
```


## See also


[FillFormat Object](Excel.FillFormat.md)

