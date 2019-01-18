---
title: Options.MatchFuzzyByte property (Word)
keywords: vbawd10.chm162988342
f1_keywords:
- vbawd10.chm162988342
ms.prod: word
api_name:
- Word.Options.MatchFuzzyByte
ms.assetid: 978d49df-a417-11b8-069e-1147067cd1ed
ms.date: 06/08/2017
localization_priority: Normal
---


# Options.MatchFuzzyByte property (Word)

 **True** if Microsoft Word ignores the distinction between full-width and half-width characters (Latin or Japanese) during a search. Read/write **Boolean**.


## Syntax

 _expression_. `MatchFuzzyByte`

 _expression_ An expression that returns an '[Options](Word.Options.md)' object.


## Example

This example sets Microsoft Word to ignore the distinction between full-width and half-width characters (Latin or Japanese) during a search.


```vb
Options.MatchFuzzyByte = True
```


## See also


[Options Object](Word.Options.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]