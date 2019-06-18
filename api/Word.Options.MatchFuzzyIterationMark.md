---
title: Options.MatchFuzzyIterationMark property (Word)
keywords: vbawd10.chm162988346
f1_keywords:
- vbawd10.chm162988346
ms.prod: word
api_name:
- Word.Options.MatchFuzzyIterationMark
ms.assetid: 24635dfe-e48a-11b7-f8fd-a8058e31e615
ms.date: 06/08/2017
localization_priority: Normal
---


# Options.MatchFuzzyIterationMark property (Word)

 **True** if Microsoft Word ignores the distinction between types of repetition marks during a search. Read/write **Boolean**.


## Syntax

_expression_. `MatchFuzzyIterationMark`

 _expression_ An expression that returns an **[Options](Word.Options.md)** object.


## Example

This example sets Microsoft Word to ignore the distinction between types of repetition marks during a search.


```vb
Options.MatchFuzzyIterationMark = True
```


## See also


[Options Object](Word.Options.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]