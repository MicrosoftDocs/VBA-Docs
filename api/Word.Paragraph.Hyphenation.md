---
title: Paragraph.Hyphenation property (Word)
keywords: vbawd10.chm156696689
f1_keywords:
- vbawd10.chm156696689
ms.prod: word
api_name:
- Word.Paragraph.Hyphenation
ms.assetid: 984aa078-9b18-7b96-d2d6-0cd603719c6b
ms.date: 06/08/2017
localization_priority: Normal
---


# Paragraph.Hyphenation property (Word)

 **True** if the specified paragraphs are included in automatic hyphenation. **False** if the specified paragraphs are to be excluded from automatic hyphenation. Read/write **Long**.


## Syntax

_expression_. `Hyphenation`

_expression_ A variable that represents a '[Paragraph](Word.Paragraph.md)' object.


## Remarks

This property can be  **True**, **False** or **wdUndefined**.


## Example

This example turns off automatic hyphenation for the first paragraph in the active document.


```vb
ActiveDocument.Paragraphs(1).Hyphenation = False
```


## See also


[Paragraph Object](Word.Paragraph.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]