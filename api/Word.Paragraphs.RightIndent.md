---
title: Paragraphs.RightIndent property (Word)
keywords: vbawd10.chm156762218
f1_keywords:
- vbawd10.chm156762218
ms.prod: word
api_name:
- Word.Paragraphs.RightIndent
ms.assetid: da5f408c-add9-05a6-bd3d-cd507af48312
ms.date: 06/08/2017
localization_priority: Normal
---


# Paragraphs.RightIndent property (Word)

Returns or sets the right indent (in points) for the specified paragraphs. Read/write  **Single**.


## Syntax

_expression_. `RightIndent`

_expression_ Required. A variable that represents a '[Paragraphs](Word.paragraphs.md)' collection.


## Example

This example sets the right indent for all paragraphs in the active document to 1 inch from the right margin. The **InchesToPoints** method is used to convert inches to points.


```vb
ActiveDocument.Paragraphs.RightIndent = InchesToPoints(1)
```


## See also


[Paragraphs Collection Object](Word.paragraphs.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]