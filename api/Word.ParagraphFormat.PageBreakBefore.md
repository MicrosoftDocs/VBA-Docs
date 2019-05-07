---
title: ParagraphFormat.PageBreakBefore property (Word)
keywords: vbawd10.chm156434536
f1_keywords:
- vbawd10.chm156434536
ms.prod: word
api_name:
- Word.ParagraphFormat.PageBreakBefore
ms.assetid: b024b5a6-4207-c490-97a6-a5eb2903c90e
ms.date: 06/08/2017
localization_priority: Normal
---


# ParagraphFormat.PageBreakBefore property (Word)

 **True** if a page break is forced before the specified paragraphs. Can be **True**, **False**, or **wdUndefined**. Read/write **Long**.


## Syntax

_expression_. `PageBreakBefore`

_expression_ A variable that represents a '[ParagraphFormat](Word.ParagraphFormat.md)' object.


## Example

This example forces a page break before the first paragraph in the selection.


```vb
Selection.Paragraphs(1).PageBreakBefore = True
```


## See also


[ParagraphFormat Object](Word.ParagraphFormat.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]