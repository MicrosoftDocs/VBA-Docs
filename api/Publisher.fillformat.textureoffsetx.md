---
title: FillFormat.TextureOffsetX property (Publisher)
keywords: vbapb10.chm2359573
f1_keywords:
- vbapb10.chm2359573
ms.prod: publisher
ms.assetid: 8023af14-0155-0387-9af7-5f7a8ea557b4
ms.date: 06/07/2019
localization_priority: Normal
---


# FillFormat.TextureOffsetX property (Publisher)

Returns or sets a **Long** that specifies the horizontal offset of the texture from the origin in [points](../language/glossary/vbe-glossary.md#point). Read/write.


## Syntax

_expression_.**TextureOffsetX**

_expression_ A variable that represents a **[FillFormat](publisher.fillformat.md)** object.


## Property value

FLOAT


## Remarks

The position of the origin is determined by the setting of the **[TextureAlignment](Publisher.fillformat.texturealignment.md)** property.

The setting of the **TextureOffsetX** property corresponds to the setting of the **Offset X** box in the **Fill** pane of the **Format Shape** dialog box in the Publisher user interface (under **Drawing Tools**, on the **Format** tab, choose **Shape Fill**, point to **Texture**, and then choose **More Textures**).



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]