---
title: Selection.DetectLanguage method (Word)
keywords: vbawd10.chm158663191
f1_keywords:
- vbawd10.chm158663191
ms.prod: word
api_name:
- Word.Selection.DetectLanguage
ms.assetid: cfbc0d54-bb00-2bd0-ad9a-e646fdcbfe46
ms.date: 06/08/2017
localization_priority: Normal
---


# Selection.DetectLanguage method (Word)

Analyzes the specified text to determine the language that it is written in.


## Syntax

_expression_. `DetectLanguage`

_expression_ Required. A variable that represents a **[Selection](Word.Selection.md)** object.


## Remarks

The results of the  **DetectLanguage** method are stored in the **LanguageID** property on a character-by-character basis. To read the **[LanguageID](Word.Language.ID.md)** property, you must first specify a selection or range of text.

If a selection contains a partial sentence, the selection is extended to the end of the sentence.

If the  **DetectLanguage** method has already been applied to the specified text, the **LanguageDetected** property is set to **True**. To reevaluate the language of the specified text, you must first set the **[LanguageDetected](Word.Document.LanguageDetected.md)** property to **False**.


## See also


[Selection Object](Word.Selection.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]