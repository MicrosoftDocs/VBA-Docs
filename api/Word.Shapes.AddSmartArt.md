---
title: Shapes.AddSmartArt method (Word)
keywords: vbawd10.chm161415196
f1_keywords:
- vbawd10.chm161415196
ms.prod: word
api_name:
- Word.Shapes.AddSmartArt
ms.assetid: 45fabbc8-eb61-2f5f-4f69-560fe1ad188a
ms.date: 06/08/2017
localization_priority: Normal
---


# Shapes.AddSmartArt method (Word)

Inserts the specified SmartArt graphic into the active document.


## Syntax

_expression_. `AddSmartArt`( `_Layout_` , `_Left_` , `_Top_` , `_Width_` , `_Height_` , `_Anchor_` )

 _expression_ An expression that returns a '[Shapes](Word.shapes.md)' object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Layout_|Required| **[SMARTARTLAYOUT]**|A [SmartArtLayout](Office.SmartArtLayout.md) object that specifies the layout for the SmartArt graphic.|
| _Left_|Optional| **Variant**|The distance, in [points](../language/glossary/vbe-glossary.md#point), from the left edge of the slide to the left edge of the SmartArt graphic.|
| _Top_|Optional| **Variant**|The distance, in [points](../language/glossary/vbe-glossary.md#point), from the top edge of the slide to the top edge of the SmartArt graphic.|
| _Width_|Optional| **Variant**|The width of the SmartArt graphic.|
| _Height_|Optional| **Variant**|The height of the SmartArt graphic.|
| _Anchor_|Optional| **Variant**|A [Range](Word.Range.md) object that represents the text to which the SmartArt graphic is bound. If _Anchor_ is specified, the anchor is positioned at the beginning of the first paragraph in the anchoring range. If this argument is omitted, the anchoring range is selected automatically and the SmartArt graphic is positioned relative to the top and left edges of the page.|

## Return value

[Shape](Word.Shape.md)


## See also


[Shapes Collection Object](Word.shapes.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]