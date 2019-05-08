---
title: Paragraphs.PageBreakBefore property (Word)
keywords: vbawd10.chm156762216
f1_keywords:
- vbawd10.chm156762216
ms.prod: word
api_name:
- Word.Paragraphs.PageBreakBefore
ms.assetid: 573ff2bc-e9df-8a6e-49eb-0773e578969d
ms.date: 06/08/2017
localization_priority: Normal
---


# Paragraphs.PageBreakBefore property (Word)

 **True** if a page break is forced before the specified paragraphs. Can be **True**, **False**, or **wdUndefined**. Read/write **Long**.


## Syntax

_expression_. `PageBreakBefore`

_expression_ A variable that represents a '[Paragraphs](Word.paragraphs.md)' collection.


## Example

This example forces a page break before the first paragraph in the selection.


```vb
Selection.Paragraphs.PageBreakBefore = True
```


## See also


[Paragraphs Collection Object](Word.paragraphs.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]