---
title: Selection.ClearCharacterDirectFormatting Method (Word)
keywords: vbawd10.chm158663689
f1_keywords:
- vbawd10.chm158663689
ms.prod: word
api_name:
- Word.Selection.ClearCharacterDirectFormatting
ms.assetid: d2138876-c832-2407-a53e-5bd4af2421b7
ms.date: 06/08/2017
---


# Selection.ClearCharacterDirectFormatting Method (Word)

Removes character formatting (formatting that has been applied manually using the buttons on the ribbon or through the dialog boxes) from the selected text.


## Syntax

 _expression_ . **ClearCharacterDirectFormatting**

 _expression_ An expression that returns a **[Selection](Word.Selection.md)** object.


## Remarks

This method does not remove character formatting that has been applied by using a character style. To remove character formatting that the user has applied by using character styles, use the  **[ClearCharacterStyle](Word.Selection.ClearCharacterStyle.md)** method. To remove all character formatting, regardless of whether the user has applied it by using character styles or manually through the formatting features in Microsoft Word, use the **[ClearCharacterAllFormatting](Word.Selection.ClearCharacterAllFormatting.md)** method.


 **Note**  For more information about how to remove paragraph formatting, see the  **[ClearParagraphAllFormatting](Word.Selection.ClearParagraphAllFormatting.md)** , **[ClearParagraphDirectFormatting](Word.Selection.ClearParagraphDirectFormatting.md)** , or **[ClearParagraphStyle](Word.Selection.ClearParagraphStyle.md)** method.


## See also


#### Concepts


[Selection Object](Word.Selection.md)

