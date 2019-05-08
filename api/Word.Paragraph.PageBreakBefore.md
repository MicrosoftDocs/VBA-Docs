---
title: Paragraph.PageBreakBefore property (Word)
keywords: vbawd10.chm156696680
f1_keywords:
- vbawd10.chm156696680
ms.prod: word
api_name:
- Word.Paragraph.PageBreakBefore
ms.assetid: 7ef33946-d598-4de1-99d8-6a045c1bbb2a
ms.date: 06/08/2017
localization_priority: Normal
---


# Paragraph.PageBreakBefore property (Word)

 **True** if a page break is forced before the specified paragraphs. Read/write **Long**.


## Syntax

_expression_. `PageBreakBefore`

_expression_ A variable that represents a '[Paragraph](Word.Paragraph.md)' object.


## Remarks

This property can be  **True**, **False**, or **wdUndefined**.


## Example

This example forces a page break before the first paragraph in the selection.


```vb
Selection.Paragraphs(1).PageBreakBefore = True
```


## See also


[Paragraph Object](Word.Paragraph.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]