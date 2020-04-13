---
title: Selection.ClearParagraphStyle method (Word)
keywords: vbawd10.chm158663686
f1_keywords:
- vbawd10.chm158663686
ms.prod: word
api_name:
- Word.Selection.ClearParagraphStyle
ms.assetid: cfbafeac-99e1-5fae-a9a0-8cf8836add94
ms.date: 06/08/2017
localization_priority: Normal
---


# Selection.ClearParagraphStyle method (Word)

Removes paragraph formatting that has been applied through paragraph styles from the selected text.


## Syntax

_expression_. `ClearParagraphStyle`

 _expression_ An expression that returns a **[Selection](Word.Selection.md)** object.


## Remarks

This method does not remove paragraph formatting that a user has applied manually. To remove manually applied paragraph formatting, use the **[ClearParagraphDirectFormatting](Word.Selection.ClearParagraphDirectFormatting.md)** method. To remove all paragraph formatting, both style and manual formatting, use the **[ClearParagraphAllFormatting](Word.Selection.ClearParagraphAllFormatting.md)** method.


> [!NOTE] 
> To remove character formatting, see the **[ClearCharacterAllFormatting](Word.Selection.ClearCharacterAllFormatting.md)**, **[ClearCharacterDirectFormatting](Word.Selection.ClearCharacterDirectFormatting.md)**, or **[ClearCharacterStyle](Word.Selection.ClearCharacterStyle.md)** method.


## See also


[Selection Object](Word.Selection.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]