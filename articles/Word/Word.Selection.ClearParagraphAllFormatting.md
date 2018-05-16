---
title: Selection.ClearParagraphAllFormatting Method (Word)
keywords: vbawd10.chm158663695
f1_keywords:
- vbawd10.chm158663695
ms.prod: word
api_name:
- Word.Selection.ClearParagraphAllFormatting
ms.assetid: b3a88322-933a-ff14-e788-e1934aba243d
ms.date: 06/08/2017
---


# Selection.ClearParagraphAllFormatting Method (Word)

Removes all paragraph formatting (formatting applied either through paragraph styles or manually applied formatting) from the selected text.


## Syntax

 _expression_ . **ClearParagraphAllFormatting**

 _expression_ An expression that returns a **[Selection](Word.Selection.md)** object.


## Remarks

This method removes all paragraph formatting. If you need to remove paragraph formatting applied through paragraph styles, use the  **[ClearParagraphStyle](Word.Selection.ClearParagraphStyle.md)** method. To remove paragraph formatting that the user has manually applied using Microsoft Word paragraph formatting features, use the **[ClearParagraphDirectFormatting](Word.Selection.ClearParagraphDirectFormatting.md)** method.


 **Note**  To remove character formatting, see the  **[ClearCharacterAllFormatting](Word.Selection.ClearCharacterAllFormatting.md)** , **[ClearCharacterDirectFormatting](Word.Selection.ClearCharacterDirectFormatting.md)** , or **[ClearCharacterStyle](Word.Selection.ClearCharacterStyle.md)** method.


## See also


#### Concepts


[Selection Object](Word.Selection.md)

