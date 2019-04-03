---
title: Find.MatchKashida property (Word)
keywords: vbawd10.chm162529380
f1_keywords:
- vbawd10.chm162529380
ms.prod: word
api_name:
- Word.Find.MatchKashida
ms.assetid: 0806a135-2238-e33e-8d0f-b0788b40754c
ms.date: 06/08/2017
localization_priority: Normal
---


# Find.MatchKashida property (Word)

 **True** if find operations match text with matching kashidas in an Arabic language document. Read/write **Boolean**.


## Syntax

_expression_. `MatchKashida`

 _expression_ An expression that returns a '[Find](Word.Find.md)' object.


## Example

This example sets the current find operation to match kashidas.


```vb
Selection.Find.MatchKashida = True
```


## See also


[Find Object](Word.Find.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]