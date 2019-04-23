---
title: Find.MatchControl property (Word)
keywords: vbawd10.chm162529383
f1_keywords:
- vbawd10.chm162529383
ms.prod: word
api_name:
- Word.Find.MatchControl
ms.assetid: 43d76f90-5b3f-db3b-15b0-98e87d8d8bc8
ms.date: 06/08/2017
localization_priority: Normal
---


# Find.MatchControl property (Word)

 **True** if find operations match text with matching bidirectional control characters in a right-to-left language document. Read/write **Boolean**.


## Syntax

_expression_. `MatchControl`

 _expression_ An expression that returns a '[Find](Word.Find.md)' object.


## Example

This example sets the current find operation to match bidirectional control characters.


```vb
Selection.Find.MatchControl = True
```


## See also


[Find Object](Word.Find.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]