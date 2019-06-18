---
title: Options.MatchFuzzySmallKana property (Word)
keywords: vbawd10.chm162988344
f1_keywords:
- vbawd10.chm162988344
ms.prod: word
api_name:
- Word.Options.MatchFuzzySmallKana
ms.assetid: 743fdfa1-01da-32ee-22cf-c30852f382bf
ms.date: 06/08/2017
localization_priority: Normal
---


# Options.MatchFuzzySmallKana property (Word)

 **True** if Microsoft Word ignores the distinction between diphthongs and double consonants during a search. Read/write **Boolean**.


## Syntax

_expression_. `MatchFuzzySmallKana`

 _expression_ An expression that returns an **[Options](Word.Options.md)** object.


## Example

This example sets Microsoft Word to ignore the distinction between diphthongs and double consonants during a search.


```vb
Options.MatchFuzzySmallKana = True
```


## See also


[Options Object](Word.Options.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]