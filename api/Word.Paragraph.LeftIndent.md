---
title: Paragraph.LeftIndent property (Word)
keywords: vbawd10.chm156696683
f1_keywords:
- vbawd10.chm156696683
ms.prod: word
api_name:
- Word.Paragraph.LeftIndent
ms.assetid: 1e30416e-fcf2-e0cd-694a-d3608fa950f8
ms.date: 06/08/2017
localization_priority: Normal
---


# Paragraph.LeftIndent property (Word)

Returns or sets a  **Single** that represents the left indent value (in points) for the specified paragraph. Read/write.


## Syntax

_expression_. `LeftIndent`

_expression_ A variable that represents a '[Paragraph](Word.Paragraph.md)' object.


## Example

This example sets the left indent of the first paragraph in the active document to 1 inch. The **InchesToPoints** method is used to convert inches to points.


```vb
ActiveDocument.Paragraphs(1).LeftIndent = InchesToPoints(1)
```


## See also


[Paragraph Object](Word.Paragraph.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]