---
title: Options.MatchFuzzyPunctuation property (Word)
keywords: vbawd10.chm162988357
f1_keywords:
- vbawd10.chm162988357
ms.prod: word
api_name:
- Word.Options.MatchFuzzyPunctuation
ms.assetid: ea4cb188-7fd1-c7e5-e520-3f0826dc3cdd
ms.date: 06/08/2017
localization_priority: Normal
---


# Options.MatchFuzzyPunctuation property (Word)

 **True** if Microsoft Word ignores the distinction between types of punctuation marks during a search. Read/write **Boolean**.


## Syntax

_expression_. `MatchFuzzyPunctuation`

 _expression_ An expression that returns an **[Options](Word.Options.md)** object.


## Example

This example sets Microsoft Word to ignore the distinction between types of punctuation marks during a search


```vb
Options.MatchFuzzyPunctuation = True
```


## See also


[Options Object](Word.Options.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]