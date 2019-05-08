---
title: ParagraphFormat.WidowControl property (Word)
keywords: vbawd10.chm156434546
f1_keywords:
- vbawd10.chm156434546
ms.prod: word
api_name:
- Word.ParagraphFormat.WidowControl
ms.assetid: 461a8d5f-2f64-b3c4-657b-0b592c482ac0
ms.date: 06/08/2017
localization_priority: Normal
---


# ParagraphFormat.WidowControl property (Word)

 **True** if the first and last lines in the specified paragraph remain on the same page as the rest of the paragraph when Word repaginates the document. Can be **True**, **False** or **wdUndefined**. Read/write **Long**.


## Syntax

_expression_. `WidowControl`

_expression_ A variable that represents a '[ParagraphFormat](Word.ParagraphFormat.md)' object.


## Example

This example formats the paragraphs in the active document so that the first or last line in a paragraph can appear by itself at the top or bottom of a page.


```vb
ActiveDocument.Paragraphs.WidowControl = False
```


## See also


[ParagraphFormat Object](Word.ParagraphFormat.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]