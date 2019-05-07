---
title: Paragraph.HangingPunctuation property (Word)
keywords: vbawd10.chm156696695
f1_keywords:
- vbawd10.chm156696695
ms.prod: word
api_name:
- Word.Paragraph.HangingPunctuation
ms.assetid: 89287cb7-1b12-4fd0-4a02-b6d4dd371d70
ms.date: 06/08/2017
localization_priority: Normal
---


# Paragraph.HangingPunctuation property (Word)

 **True** if hanging punctuation is enabled for the specified paragraphs. This property returns **wdUndefined** if it's set to **True** for only some of the specified paragraphs. Read/write **Long**.


## Syntax

_expression_. `HangingPunctuation`

_expression_ A variable that represents a '[Paragraph](Word.Paragraph.md)' object.


## Example

This example enables hanging punctuation for the first paragraph in the active document.


```vb
ActiveDocument.Paragraphs(1).HangingPunctuation = True
```


## See also


[Paragraph Object](Word.Paragraph.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]