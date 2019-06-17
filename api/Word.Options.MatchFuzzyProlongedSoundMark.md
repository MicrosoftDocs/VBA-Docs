---
title: Options.MatchFuzzyProlongedSoundMark property (Word)
keywords: vbawd10.chm162988349
f1_keywords:
- vbawd10.chm162988349
ms.prod: word
api_name:
- Word.Options.MatchFuzzyProlongedSoundMark
ms.assetid: ec86cda2-3002-ff44-7657-bb70f1bf1a79
ms.date: 06/08/2017
localization_priority: Normal
---


# Options.MatchFuzzyProlongedSoundMark property (Word)

 **True** if Microsoft Word ignores the distinction between short and long vowel sounds during a search. Read/write **Boolean**.


## Syntax

_expression_. `MatchFuzzyProlongedSoundMark`

 _expression_ An expression that returns an **[Options](Word.Options.md)** object.


## Example

This example sets Microsoft Word to ignore the distinction between short and long vowel sounds during a search.


```vb
Options.MatchFuzzyProlongedSoundMark = True
```


## See also


[Options Object](Word.Options.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]