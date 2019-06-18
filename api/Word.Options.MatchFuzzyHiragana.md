---
title: Options.MatchFuzzyHiragana property (Word)
keywords: vbawd10.chm162988343
f1_keywords:
- vbawd10.chm162988343
ms.prod: word
api_name:
- Word.Options.MatchFuzzyHiragana
ms.assetid: 772b8dd9-f4be-75f4-d9ac-cbe00922d3fa
ms.date: 06/08/2017
localization_priority: Normal
---


# Options.MatchFuzzyHiragana property (Word)

 **True** if Microsoft Word ignores the distinction between hiragana and katakana during a search. Read/write **Boolean**.


## Syntax

_expression_. `MatchFuzzyHiragana`

 _expression_ An expression that returns an **[Options](Word.Options.md)** object.


## Example

This example sets Microsoft Word to ignore the distinction between hiragana and katakana during a search.


```vb
Options.MatchFuzzyHiragana = True
```


## See also


[Options Object](Word.Options.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]