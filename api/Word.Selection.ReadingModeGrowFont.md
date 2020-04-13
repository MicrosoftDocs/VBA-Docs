---
title: Selection.ReadingModeGrowFont method (Word)
keywords: vbawd10.chm158663693
f1_keywords:
- vbawd10.chm158663693
ms.prod: word
api_name:
- Word.Selection.ReadingModeGrowFont
ms.assetid: 5a23b50e-073f-1cbd-e1df-6ee846cb1ecf
ms.date: 06/08/2017
localization_priority: Normal
---


# Selection.ReadingModeGrowFont method (Word)

Increases the size of the displayed text one point size when the document is displayed in Reading mode.


## Syntax

_expression_. `ReadingModeGrowFont`

 _expression_ An expression that returns a [Selection](./Word.Selection.md) object.


## Return value

Nothing


## Remarks

Use the **[ReadingModeShrinkFont](Word.Selection.ReadingModeShrinkFont.md)** method to decrease the size of the text. This does not affect the size of the font in the document, only the size of the text while viewing the document in Reading mode.


> [!NOTE] 
> Text does not need to be selected for this method to affect the text displayed in Reading mode. Text size for all text displayed in Reading mode is affected.


## See also


[Selection Object](Word.Selection.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]