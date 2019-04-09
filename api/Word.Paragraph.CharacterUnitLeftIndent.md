---
title: Paragraph.CharacterUnitLeftIndent property (Word)
keywords: vbawd10.chm156696703
f1_keywords:
- vbawd10.chm156696703
ms.prod: word
api_name:
- Word.Paragraph.CharacterUnitLeftIndent
ms.assetid: 1dbe6053-52fd-f17c-aa95-3cfdef1222d5
ms.date: 06/08/2017
localization_priority: Normal
---


# Paragraph.CharacterUnitLeftIndent property (Word)

Returns or sets the left indent value (in characters) for the specified paragraphs. Read/write  **Single**.


## Syntax

_expression_. `CharacterUnitLeftIndent`

_expression_ Required. A variable that represents a '[Paragraph](Word.Paragraph.md)' object.


## Example

This example sets the left indent of the first paragraph in the active document to one character from the left margin.


```vb
ActiveDocument.Paragraphs(1) _ 
 .CharacterUnitLeftIndent = 1
```


## See also


[Paragraph Object](Word.Paragraph.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]