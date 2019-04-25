---
title: FillFormat.TextureType property (Excel)
keywords: vbaxl10.chm115021
f1_keywords:
- vbaxl10.chm115021
ms.prod: excel
api_name:
- Excel.FillFormat.TextureType
ms.assetid: 9a39c34e-c19c-5539-b5ac-b624fe71e2e9
ms.date: 04/26/2019
localization_priority: Normal
---


# FillFormat.TextureType property (Excel)

Returns the texture type for the specified fill. Read-only **[MsoTextureType](Office.MsoTextureType.md)**.


## Syntax

_expression_.**TextureType**

_expression_ A variable that represents a **[FillFormat](Excel.FillFormat.md)** object.


## Remarks

Use the **[UserTextured](Excel.FillFormat.UserTextured.md)** method to set the texture type for the fill.


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




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]