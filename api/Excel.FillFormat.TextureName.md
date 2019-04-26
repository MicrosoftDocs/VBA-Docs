---
title: FillFormat.TextureName property (Excel)
keywords: vbaxl10.chm115020
f1_keywords:
- vbaxl10.chm115020
ms.prod: excel
api_name:
- Excel.FillFormat.TextureName
ms.assetid: 9ef98f75-6407-010c-5c8f-44f3d236c04f
ms.date: 04/26/2019
localization_priority: Normal
---


# FillFormat.TextureName property (Excel)

Returns the name of the custom texture file for the specified fill. Read-only **String**.


## Syntax

_expression_.**TextureName**

_expression_ A variable that represents a **[FillFormat](Excel.FillFormat.md)** object.


## Remarks

Use the **[UserPicture](Excel.FillFormat.UserPicture.md)** or **[UserTextured](Excel.FillFormat.UserTextured.md)** method to set the texture file for the fill.


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