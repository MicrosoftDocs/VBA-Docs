---
title: Options.MatchFuzzyKanji property (Word)
keywords: vbawd10.chm162988347
f1_keywords:
- vbawd10.chm162988347
ms.prod: word
api_name:
- Word.Options.MatchFuzzyKanji
ms.assetid: 6d2a1b1f-2a1c-23d2-5e3b-aa8f2e26388e
ms.date: 06/08/2017
localization_priority: Normal
---


# Options.MatchFuzzyKanji property (Word)

 **True** if Microsoft Word ignores the distinction between standard and nonstandard kanji ideography during a search. Read/write **Boolean**.


## Syntax

_expression_. `MatchFuzzyKanji`

 _expression_ An expression that returns an **[Options](Word.Options.md)** object.


## Example

This example sets Microsoft Word to ignore the distinction between standard and nonstandard Kanji ideography during a search.


```vb
Options.MatchFuzzyKanji = True
```


## See also


[Options Object](Word.Options.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]