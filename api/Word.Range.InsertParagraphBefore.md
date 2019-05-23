---
title: Range.InsertParagraphBefore method (Word)
keywords: vbawd10.chm157155540
f1_keywords:
- vbawd10.chm157155540
ms.prod: word
api_name:
- Word.Range.InsertParagraphBefore
ms.assetid: 78d62099-fa2c-911d-690b-93a9ee4f58eb
ms.date: 06/08/2017
localization_priority: Normal
---


# Range.InsertParagraphBefore method (Word)

Inserts a new paragraph before the specified range.


## Syntax

_expression_. `InsertParagraphBefore`

_expression_ Required. A variable that represents a **[Range](Word.Range.md)** object.


## Remarks

After this method is applied, the range expands to include the new paragraph.


## Example

This example inserts a new paragraph at the beginning of the active document.


```vb
ActiveDocument.Range.InsertParagraphBefore
```


## See also


[Range Object](Word.Range.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]