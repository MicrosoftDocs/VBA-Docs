---
title: Paragraph.Reset method (Word)
keywords: vbawd10.chm156696888
f1_keywords:
- vbawd10.chm156696888
ms.prod: word
api_name:
- Word.Paragraph.Reset
ms.assetid: 9a2ac15e-406e-2e83-114c-82fa2324f26a
ms.date: 06/08/2017
localization_priority: Normal
---


# Paragraph.Reset method (Word)

Removes manual paragraph formatting (formatting not applied using a style).


## Syntax

_expression_. `Reset`

_expression_ Required. A variable that represents a '[Paragraph](Word.Paragraph.md)' object.


## Remarks

If you manually right align a paragraph and the underlying style has a different alignment, the **Reset** method changes the alignment to match the formatting of the underlying style.


## Example

This example removes manual paragraph formatting from the second paragraph in the active document.


```vb
ActiveDocument.Paragraphs(2).Reset
```


## See also


[Paragraph Object](Word.Paragraph.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]