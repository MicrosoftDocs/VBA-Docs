---
title: Font.Ligature property (Publisher)
keywords: vbapb10.chm5374007
f1_keywords:
- vbapb10.chm5374007
ms.prod: publisher
ms.assetid: 17847824-8761-42b7-8d0c-00345e8b5de8
ms.date: 06/08/2019
localization_priority: Normal
---


# Font.Ligature property (Publisher)

Returns or sets a **[PbLigaturePresetType](Publisher.pbligaturepresettype.md)** constant that represents the state of the **Ligature** property on the characters in a text range. The **Ligature** property enables embellishments to the characters, often in the form of bigger and more flamboyant serifs. Read/write.


## Syntax

_expression_.**Ligature**

_expression_ A variable that represents a **[Font](Publisher.Font.md)** object.


## Return value

PbLigaturePresetType


## Remarks

The **Ligature** property has an effect only for OpenType fonts that contain ligatures.

Ligatures are alternate appearances of sequences of characters; multiple characters are merged into one glyph. For example, when ligatures are turned on for the word _Office_, the letters _ffi_ are all joined together into one glyph that displays a continuous line from the first _f_ through the dot in the _i_.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]