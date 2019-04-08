---
title: Replacement.LanguageIDFarEast property (Word)
keywords: vbawd10.chm162594835
f1_keywords:
- vbawd10.chm162594835
ms.prod: word
api_name:
- Word.Replacement.LanguageIDFarEast
ms.assetid: 66029c49-d297-5685-525c-79d7cacae1af
ms.date: 06/08/2017
localization_priority: Normal
---


# Replacement.LanguageIDFarEast property (Word)

Returns or sets an East Asian language for the specified replacement. Read/write  **[WdLanguageID](Word.WdLanguageID.md)**.


## Syntax

_expression_. `LanguageIDFarEast`

_expression_ Required. A variable that represents a '[Replacement](Word.Replacement.md)' object.


## Remarks

This is the recommended way to return or set the language of East Asian text in a document created in an East Asian version of Microsoft Word.


## Example

This example sets the language of the selection to Korean.


```vb
Selection.LanguageIDFarEast = wdKorean
```


## See also


[Replacement Object](Word.Replacement.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]