---
title: InlineShapes.AddSmartArt method (Word)
keywords: vbawd10.chm162070636
f1_keywords:
- vbawd10.chm162070636
ms.prod: word
api_name:
- Word.InlineShapes.AddSmartArt
ms.assetid: 7ece8207-2bb9-d88d-25c4-e2f29f3abb38
ms.date: 06/08/2017
localization_priority: Normal
---


# InlineShapes.AddSmartArt method (Word)

Inserts a SmartArt graphic as an inline shape into the active document.


## Syntax

_expression_. `AddSmartArt`( `_Layout_` , `_Range_` )

 _expression_ An expression that returns a '[InlineShapes](Word.inlineshapes.md)' object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Layout_|Required| **[SMARTARTLAYOUT]**|A [SmartArtLayout](Office.SmartArtLayout.md)object that specifies the layout for the SmartArt graphic.|
| _Range_|Optional| **Variant**|Specifies the text to which the SmartArt graphic is bound. If [Range](Word.Range.md) is specified, the SmartArt graphic is positioned at the beginning of the first paragraph in the range. If this argument is omitted, the range is selected automatically, and the SmartArt graphic is positioned relative to the top and left edges of the page.|

## Return value

InlineShape


## See also


[InlineShapes Collection Object](Word.inlineshapes.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]