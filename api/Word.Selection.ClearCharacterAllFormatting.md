---
title: Selection.ClearCharacterAllFormatting method (Word)
keywords: vbawd10.chm158663687
f1_keywords:
- vbawd10.chm158663687
ms.prod: word
api_name:
- Word.Selection.ClearCharacterAllFormatting
ms.assetid: 1d0dfb43-4855-1534-5ec2-475232a6a457
ms.date: 06/08/2017
localization_priority: Normal
---


# Selection.ClearCharacterAllFormatting method (Word)

Removes all character formatting (formatting applied either through character styles or manually applied formatting) from the selected text.


## Syntax

_expression_. `ClearCharacterAllFormatting`

 _expression_ An expression that returns a **[Selection](Word.Selection.md)** object.


## Remarks

This method removes all character formatting. If you need to removed formatting applied through character styles, use the **[ClearCharacterStyle](Word.Selection.ClearCharacterStyle.md)** method. To remove character formatting that the user has manually applied using Microsoft Word character formatting features, use the **[ClearCharacterDirectFormatting](Word.Selection.ClearCharacterDirectFormatting.md)** method.


> [!NOTE] 
> To remove paragraph formatting, see the **[ClearParagraphAllFormatting](Word.Selection.ClearParagraphAllFormatting.md)**, **[ClearParagraphDirectFormatting](Word.Selection.ClearParagraphDirectFormatting.md)**, or **[ClearParagraphStyle](Word.Selection.ClearParagraphStyle.md)** method.


## See also


[Selection Object](Word.Selection.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]