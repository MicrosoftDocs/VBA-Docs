---
title: Selection.ReadingModeShrinkFont method (Word)
keywords: vbawd10.chm158663694
f1_keywords:
- vbawd10.chm158663694
ms.prod: word
api_name:
- Word.Selection.ReadingModeShrinkFont
ms.assetid: 58472c33-7f8e-dc3b-04d8-7b50ca911ed4
ms.date: 06/08/2017
localization_priority: Normal
---


# Selection.ReadingModeShrinkFont method (Word)

Decreases the size of the displayed text one point size when the document is displayed in Reading mode.


## Syntax

_expression_. `ReadingModeShrinkFont`

 _expression_ An expression that returns a [Selection](./Word.Selection.md) object.


## Return value

Nothing


## Remarks

Use the **[ReadingModeGrowFont](Word.Selection.ReadingModeGrowFont.md)** method to increase the size of the text. This does not affect the size of the font in the document, only the size of the text while viewing the document in Reading mode.


> [!NOTE] 
> Text does not need to be selected for this method to affect the text displayed in Reading mode. Text size for all text displayed in Reading mode is affected.


## See also


[Selection Object](Word.Selection.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]