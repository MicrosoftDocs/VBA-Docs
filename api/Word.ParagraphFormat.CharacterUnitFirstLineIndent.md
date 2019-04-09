---
title: ParagraphFormat.CharacterUnitFirstLineIndent property (Word)
keywords: vbawd10.chm156434560
f1_keywords:
- vbawd10.chm156434560
ms.prod: word
api_name:
- Word.ParagraphFormat.CharacterUnitFirstLineIndent
ms.assetid: f5e68ef0-7086-4d33-7ed0-3c0569d203cd
ms.date: 06/08/2017
localization_priority: Normal
---


# ParagraphFormat.CharacterUnitFirstLineIndent property (Word)

Returns or sets the value (in characters) for a first-line or hanging indent. Use a positive value to set a first-line indent, and use a negative value to set a hanging indent. Read/write  **Single**.


## Syntax

_expression_. `CharacterUnitFirstLineIndent`

_expression_ Required. A variable that represents a '[ParagraphFormat](Word.ParagraphFormat.md)' object.


## Example

This example sets a first-line indent of one character for the first paragraph in the active document.


```vb
ActiveDocument.Paragraphs(1) _ 
 .CharacterUnitFirstLineIndent = 1
```

This example sets a hanging indent of 1.5 characters for the second paragraph in the active document.




```vb
ActiveDocument.Paragraphs(2) _ 
 .CharacterUnitFirstLineIndent = -1.5
```


## See also


[ParagraphFormat Object](Word.ParagraphFormat.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]