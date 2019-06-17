---
title: Options.MatchFuzzySpace property (Word)
keywords: vbawd10.chm162988358
f1_keywords:
- vbawd10.chm162988358
ms.prod: word
api_name:
- Word.Options.MatchFuzzySpace
ms.assetid: b32a93ac-620f-ba6a-a6e9-e38d72eda5cf
ms.date: 06/08/2017
localization_priority: Normal
---


# Options.MatchFuzzySpace property (Word)

 **True** if Microsoft Word ignores the distinction between space markers used during a search. Read/write **Boolean**.


## Syntax

_expression_. `MatchFuzzySpace`

 _expression_ An expression that returns an **[Options](Word.Options.md)** object.


## Example

This example sets Microsoft Word to ignore the distinction between space markers used during a search.


```vb
Options.MatchFuzzySpace = True
```


## See also


[Options Object](Word.Options.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]