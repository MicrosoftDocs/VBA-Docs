---
title: Style.LanguageIDFarEast property (Word)
keywords: vbawd10.chm153878544
f1_keywords:
- vbawd10.chm153878544
ms.prod: word
api_name:
- Word.Style.LanguageIDFarEast
ms.assetid: f36c06a7-82e8-f934-9566-4c1275ed3e8c
ms.date: 06/08/2017
localization_priority: Normal
---


# Style.LanguageIDFarEast property (Word)

Returns or sets an East Asian language for the specified object. Read/write  **[WdLanguageID](Word.WdLanguageID.md)**.


## Syntax

_expression_. `LanguageIDFarEast`

_expression_ Required. A variable that represents a '[Style](Word.Style.md)' object.


## Remarks

This is the recommended way to return or set the language of East Asian text in a document created in an East Asian version of Microsoft Word.


## Example

This example sets the language of the selection to Korean.


```vb
Selection.LanguageIDFarEast = wdKorean
```


## See also


[Style Object](Word.Style.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]