---
title: FillFormat.TextureOffsetY property (Publisher)
keywords: vbapb10.chm2359574
f1_keywords:
- vbapb10.chm2359574
ms.prod: publisher
ms.assetid: aa690d54-a4b1-5073-1957-13a638cf3e19
ms.date: 06/07/2019
localization_priority: Normal
---


# FillFormat.TextureOffsetY property (Publisher)

Returns or sets a **Long** that specifies the vertical offset of the texture from the origin in [points](../language/glossary/vbe-glossary.md#point). Read/write.


## Syntax

_expression_.**TextureOffsetY**

_expression_ A variable that represents a **[FillFormat](publisher.fillformat.md)** object.


## Property value

FLOAT


## Remarks

The position of the origin is determined by the setting of the **[TextureAlignment](Publisher.fillformat.texturealignment.md)** property.

The setting of the **TextureOffsetY** property corresponds to the setting of the **Offset Y** box in the **Fill** pane of the **Format Shape** dialog box in the Publisher user interface (under **Drawing Tools**, on the **Format** tab, choose **Shape Fill**, point to **Texture**, and then choose **More Textures**).



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]