---
title: Paragraphs.LeftIndent property (Word)
keywords: vbawd10.chm156762219
f1_keywords:
- vbawd10.chm156762219
ms.prod: word
api_name:
- Word.Paragraphs.LeftIndent
ms.assetid: 543bfc55-77c1-3db3-ed61-b5c8cdb7cae0
ms.date: 06/08/2017
localization_priority: Normal
---


# Paragraphs.LeftIndent property (Word)

Returns or sets a  **Single** that represents the left indent value (in points) for the specified paragraphs. Read/write.


## Syntax

_expression_. `LeftIndent`

_expression_ A variable that represents a '[Paragraphs](Word.paragraphs.md)' collection.


## Example

This example sets the left indent of all paragraphs in the active document to 1 inch. The  **InchesToPoints** method is used to convert inches to points.


```vb
ActiveDocument.Paragraphs.LeftIndent = InchesToPoints(1)
```


## See also


[Paragraphs Collection Object](Word.paragraphs.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]