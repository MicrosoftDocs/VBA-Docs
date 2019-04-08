---
title: Paragraph.CharacterUnitRightIndent property (Word)
keywords: vbawd10.chm156696702
f1_keywords:
- vbawd10.chm156696702
ms.prod: word
api_name:
- Word.Paragraph.CharacterUnitRightIndent
ms.assetid: f7241ec4-7737-3393-9a78-45a2dd267b8f
ms.date: 06/08/2017
localization_priority: Normal
---


# Paragraph.CharacterUnitRightIndent property (Word)

Returns or sets the right indent value (in characters) for the specified paragraphs. Read/write  **Single**.


## Syntax

_expression_. `CharacterUnitRightIndent`

_expression_ Required. A variable that represents a '[Paragraph](Word.Paragraph.md)' object.


## Example

This example sets the right indent for all paragraphs in the active document to one character from the right margin.


```vb
ActiveDocument.Paragraphs _ 
 .CharacterUnitRightIndent = 1
```


## See also


[Paragraph Object](Word.Paragraph.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]