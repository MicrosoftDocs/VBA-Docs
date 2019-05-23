---
title: Selection.LanguageIDFarEast property (Word)
keywords: vbawd10.chm158662810
f1_keywords:
- vbawd10.chm158662810
ms.prod: word
api_name:
- Word.Selection.LanguageIDFarEast
ms.assetid: 59f5b72f-3ba5-cff8-8465-6759d2194d26
ms.date: 06/08/2017
localization_priority: Normal
---


# Selection.LanguageIDFarEast property (Word)

Returns or sets an East Asian language for the specified object. Read/write  **WdLanguageID**.


## Syntax

_expression_. `LanguageIDFarEast`

_expression_ Required. A variable that represents a **[Selection](Word.Selection.md)** object.


## Remarks

This is the recommended way to return or set the language of East Asian text in a document created in an East Asian version of Microsoft Word.


## Example

This example sets the language of the selection to Korean.


```vb
Selection.LanguageIDFarEast = wdKorean
```


## See also


[Selection Object](Word.Selection.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]