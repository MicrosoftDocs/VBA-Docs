---
title: Selection.ClearCharacterStyle method (Word)
keywords: vbawd10.chm158663688
f1_keywords:
- vbawd10.chm158663688
ms.prod: word
api_name:
- Word.Selection.ClearCharacterStyle
ms.assetid: ff9795f9-ea74-fa03-5d87-9c56152d179d
ms.date: 06/08/2017
localization_priority: Normal
---


# Selection.ClearCharacterStyle method (Word)

Removes character formatting that has been applied through character styles from the selected text.


## Syntax

_expression_. `ClearCharacterStyle`

 _expression_ An expression that returns a **[Selection](Word.Selection.md)** object.


## Remarks

This method does not remove character formatting that a user has applied manually. To remove manually applied character formatting, use the  **[ClearCharacterDirectFormatting](Word.Selection.ClearCharacterDirectFormatting.md)** method. To remove all character formatting, both style and manual formatting, use the **[ClearCharacterAllFormatting](Word.Selection.ClearCharacterAllFormatting.md)** method.


> [!NOTE] 
> To remove paragraph formatting, see the  **[ClearParagraphAllFormatting](Word.Selection.ClearParagraphAllFormatting.md)**, **[ClearParagraphDirectFormatting](Word.Selection.ClearParagraphDirectFormatting.md)**, or **[ClearParagraphStyle](Word.Selection.ClearParagraphStyle.md)** method.


## See also


[Selection Object](Word.Selection.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]