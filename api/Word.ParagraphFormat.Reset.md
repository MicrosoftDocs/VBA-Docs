---
title: ParagraphFormat.Reset method (Word)
keywords: vbawd10.chm156434744
f1_keywords:
- vbawd10.chm156434744
ms.prod: word
api_name:
- Word.ParagraphFormat.Reset
ms.assetid: ba44a672-1a02-e673-9bee-b0a7239445a2
ms.date: 06/08/2017
localization_priority: Normal
---


# ParagraphFormat.Reset method (Word)

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


[ParagraphFormat Object](Word.ParagraphFormat.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]