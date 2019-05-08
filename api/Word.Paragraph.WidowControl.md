---
title: Paragraph.WidowControl property (Word)
keywords: vbawd10.chm156696690
f1_keywords:
- vbawd10.chm156696690
ms.prod: word
api_name:
- Word.Paragraph.WidowControl
ms.assetid: 5bf158e5-02e4-03f8-0f48-c596d53dc13a
ms.date: 06/08/2017
localization_priority: Normal
---


# Paragraph.WidowControl property (Word)

 **True** if the first and last lines in the specified paragraph remain on the same page as the rest of the paragraph when Word repaginates the document. Read/write **Long**.


## Syntax

_expression_. `WidowControl`

_expression_ A variable that represents a '[Paragraph](Word.Paragraph.md)' object.


## Remarks

This property can be  **True**, **False** or **wdUndefined**.


## Example

This example formats the first paragraph in the active document so that the first or last line in a paragraph can appear by itself at the top or bottom of a page.


```vb
ActiveDocument.Paragraphs(1).WidowControl = False
```


## See also


[Paragraph Object](Word.Paragraph.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]