---
title: ParagraphFormat.RightIndent property (Word)
keywords: vbawd10.chm156434538
f1_keywords:
- vbawd10.chm156434538
ms.prod: word
api_name:
- Word.ParagraphFormat.RightIndent
ms.assetid: de69209e-d88d-d367-9d84-94faa07a30bd
ms.date: 06/08/2017
localization_priority: Normal
---


# ParagraphFormat.RightIndent property (Word)

Returns or sets the right indent (in points) for the specified paragraphs. Read/write  **Single**.


## Syntax

_expression_. `RightIndent`

_expression_ Required. A variable that represents a '[ParagraphFormat](Word.ParagraphFormat.md)' object.


## Example

This example sets the right indent for all paragraphs in the active document to 1 inch from the right margin. The **InchesToPoints** method is used to convert inches to points.


```vb
ActiveDocument.Paragraphs.RightIndent = InchesToPoints(1)
```


## See also


[ParagraphFormat Object](Word.ParagraphFormat.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]