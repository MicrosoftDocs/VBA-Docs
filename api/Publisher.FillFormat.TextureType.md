---
title: FillFormat.TextureType property (Publisher)
keywords: vbapb10.chm2359568
f1_keywords:
- vbapb10.chm2359568
ms.prod: publisher
api_name:
- Publisher.FillFormat.TextureType
ms.assetid: 08f3b0a1-97a3-bdbf-25b4-93e05938d607
ms.date: 06/07/2019
localization_priority: Normal
---


# FillFormat.TextureType property (Publisher)

Returns an **[MsoTextureType](office.msotexturetype.md)** constant indicating the texture type for the specified fill. Read-only.


## Syntax

_expression_.**TextureType**

_expression_ A variable that represents a **[FillFormat](publisher.fillformat.md)** object.


## Return value

MsoTextureType


## Remarks

This property is read-only. Use the **[PresetTextured](Publisher.FillFormat.PresetTextured.md)** or **[UserTextured](Publisher.FillFormat.UserTextured.md)** method to set the texture type for the fill.

The property value can be one of the **MsoTextureType** constants declared in the Microsoft Office type library. 

## Example

This example applies a canvas texture to the fill for all shapes on the first page of the active publication that currently have fills with a user-defined texture.

```vb
Dim shpLoop As Shape 
 
For Each shpLoop In ActiveDocument.Pages(1).Shapes 
 With shpLoop.Fill 
 If .TextureType = msoTextureUserDefined Then 
 .PresetTextured _ 
 PresetTexture:=msoTextureCanvas 
 End If 
 End With 
Next shpLoop
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]