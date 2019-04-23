---
title: Paragraph.LineUnitBefore property (Word)
keywords: vbawd10.chm156696705
f1_keywords:
- vbawd10.chm156696705
ms.prod: word
api_name:
- Word.Paragraph.LineUnitBefore
ms.assetid: e9947ad7-14aa-b261-7b2c-c26ad05863eb
ms.date: 06/08/2017
localization_priority: Normal
---


# Paragraph.LineUnitBefore property (Word)

Returns or sets the amount of spacing (in gridlines) before the specified paragraph. Read/write  **Single**.


## Syntax

_expression_. `LineUnitBefore`

_expression_ Required. A variable that represents a '[Paragraph](Word.Paragraph.md)' object.


## Example

This example sets the spacing before the second paragraph in the active document to one gridline.


```vb
ActiveDocument.Paragraphs(2).LineUnitBefore = 1
```


## See also


[Paragraph Object](Word.Paragraph.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]