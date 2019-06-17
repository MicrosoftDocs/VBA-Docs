---
title: Options.MatchFuzzyOldKana property (Word)
keywords: vbawd10.chm162988348
f1_keywords:
- vbawd10.chm162988348
ms.prod: word
api_name:
- Word.Options.MatchFuzzyOldKana
ms.assetid: 682e9473-8e0f-b5cc-1c17-4b16ee499280
ms.date: 06/08/2017
localization_priority: Normal
---


# Options.MatchFuzzyOldKana property (Word)

 **True** if Microsoft Word ignores the distinction between new kana and old kana characters during a search. Read/write **Boolean**.


## Syntax

_expression_. `MatchFuzzyOldKana`

 _expression_ An expression that returns an **[Options](Word.Options.md)** object.


## Example

This example sets Microsoft Word to ignore the distinction between new kana and old kana characters during a search.


```vb
Options.MatchFuzzyOldKana = True
```


## See also


[Options Object](Word.Options.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]