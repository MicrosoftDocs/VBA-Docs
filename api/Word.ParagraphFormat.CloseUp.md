---
title: ParagraphFormat.CloseUp method (Word)
keywords: vbawd10.chm156434733
f1_keywords:
- vbawd10.chm156434733
ms.prod: word
api_name:
- Word.ParagraphFormat.CloseUp
ms.assetid: 021ab4fe-3301-90c7-2543-59140b7881da
ms.date: 06/08/2017
localization_priority: Normal
---


# ParagraphFormat.CloseUp method (Word)

Removes any spacing before paragraphs in the specified paragraph format.


## Syntax

_expression_. `CloseUp`

_expression_ Required. A variable that represents a '[ParagraphFormat](Word.ParagraphFormat.md)' object.


## Remarks

The following two statements are equivalent: 


```vb
ParagraphFormat.CloseUp 
ParagraphFormat.SpaceBefore = 0
```


## See also


[ParagraphFormat Object](Word.ParagraphFormat.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]