---
title: Range.EmphasisMark property (Word)
keywords: vbawd10.chm157155468
f1_keywords:
- vbawd10.chm157155468
ms.prod: word
api_name:
- Word.Range.EmphasisMark
ms.assetid: 6f0f7d19-efba-8fee-7e6c-abb1defe8529
ms.date: 06/08/2017
localization_priority: Normal
---


# Range.EmphasisMark property (Word)

Returns or sets the emphasis mark for a character or designated character string. Read/write  **WdEmphasisMark**.


## Syntax

_expression_. `EmphasisMark`

_expression_ Required. A variable that represents a **[Range](Word.Range.md)** object.


## Example

This example sets the emphasis mark over the fourth word in the active document to a comma.


```vb
ActiveDocument.Words(4).EmphasisMark = wdEmphasisMarkOverComma
```


## See also


[Range Object](Word.Range.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]