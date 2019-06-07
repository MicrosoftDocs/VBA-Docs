---
title: Find.MatchAlefHamza property (Word)
keywords: vbawd10.chm162529382
f1_keywords:
- vbawd10.chm162529382
ms.prod: word
api_name:
- Word.Find.MatchAlefHamza
ms.assetid: 1023d28a-d6b7-658a-0fb2-e2f9bd11b457
ms.date: 06/08/2017
localization_priority: Normal
---


# Find.MatchAlefHamza property (Word)

**True** if find operations match text with matching alef hamzas in an Arabic language document. Read/write **Boolean**.


## Syntax

_expression_.**MatchAlefHamza**

_expression_ An expression that returns a **[Find](Word.Find.md)** object.


## Example

This example sets the current find operation to match alef hamzas.

```vb
Selection.Find.MatchAlefHamza = True
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]