---
title: Selection.ClearParagraphDirectFormatting method (Word)
keywords: vbawd10.chm158663696
f1_keywords:
- vbawd10.chm158663696
ms.prod: word
api_name:
- Word.Selection.ClearParagraphDirectFormatting
ms.assetid: 66df2319-f02e-7cd9-4cef-fda6468dcd67
ms.date: 06/08/2017
localization_priority: Normal
---


# Selection.ClearParagraphDirectFormatting method (Word)

Removes paragraph formatting that has been applied manually (using the buttons on the ribbon or through the dialog boxes) from the selected text.


## Syntax

_expression_. `ClearParagraphDirectFormatting`

 _expression_ An expression that returns a **[Selection](Word.Selection.md)** object.


## Remarks

This method does not remove paragraph formatting that has been applied by using a paragraph style. To remove paragraph formatting that the user has applied by using paragraph styles, use the  **[ClearParagraphStyle](Word.Selection.ClearParagraphStyle.md)** method. To remove all paragraph formatting, both style and manual formatting, use the **[ClearParagraphAllFormatting](Word.Selection.ClearParagraphAllFormatting.md)** method.


> [!NOTE] 
> For more information about how to remove character formatting, see the  **[ClearCharacterAllFormatting](Word.Selection.ClearCharacterAllFormatting.md)**, **[ClearCharacterDirectFormatting](Word.Selection.ClearCharacterDirectFormatting.md)**, or **[ClearCharacterStyle](Word.Selection.ClearCharacterStyle.md)** method.


## See also


[Selection Object](Word.Selection.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]