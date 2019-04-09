---
title: Paragraphs.LineUnitBefore property (Word)
keywords: vbawd10.chm156762241
f1_keywords:
- vbawd10.chm156762241
ms.prod: word
api_name:
- Word.Paragraphs.LineUnitBefore
ms.assetid: 8db3f0e4-1f52-ce37-b685-e8ace269d1d5
ms.date: 06/08/2017
localization_priority: Normal
---


# Paragraphs.LineUnitBefore property (Word)

Returns or sets the amount of spacing (in gridlines) before the specified paragraphs. Read/write  **Single**.


## Syntax

_expression_. `LineUnitBefore`

_expression_ Required. A variable that represents a '[Paragraphs](Word.paragraphs.md)' collection.


## Example

This example sets the spacing before all paragraphs in the active document to one gridline.


```vb
ActiveDocument.Paragraphs.LineUnitBefore = 1
```


## See also


[Paragraphs Collection Object](Word.paragraphs.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]