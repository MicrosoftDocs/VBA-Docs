---
title: Paragraphs.Format property (Word)
keywords: vbawd10.chm156763214
f1_keywords:
- vbawd10.chm156763214
ms.prod: word
api_name:
- Word.Paragraphs.Format
ms.assetid: 7f087836-82ad-829e-5529-258ba4a3a9b1
ms.date: 06/08/2017
localization_priority: Normal
---


# Paragraphs.Format property (Word)

Returns or sets a  **[ParagraphFormat](Word.ParagraphFormat.md)** object that represents the formatting of the specified paragraph or paragraphs.


## Syntax

_expression_.**Format**

_expression_ Required. A variable that represents a '[Paragraphs](Word.paragraphs.md)' collection.


## Example

The following example left-aligns all the paragraphs in the active document.


```vb
ActiveDocument.Paragraphs.Format.Alignment = wdAlignParagraphLeft
```


## See also


[Paragraphs Collection Object](Word.paragraphs.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]