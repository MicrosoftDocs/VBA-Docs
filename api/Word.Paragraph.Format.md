---
title: Paragraph.Format property (Word)
keywords: vbawd10.chm156697678
f1_keywords:
- vbawd10.chm156697678
ms.prod: word
api_name:
- Word.Paragraph.Format
ms.assetid: d8787b8e-54c7-1adf-75b3-de7081fdff8d
ms.date: 06/08/2017
localization_priority: Normal
---


# Paragraph.Format property (Word)

Returns or sets a  **[ParagraphFormat](Word.ParagraphFormat.md)** object that represents the formatting of the specified paragraph or paragraphs.


## Syntax

_expression_.**Format**

_expression_ Required. A variable that represents a '[Paragraph](Word.Paragraph.md)' object.


## Example

This example returns the formatting of the first paragraph in the active document and then applies the formatting to the selection.


```vb
Set paraFormat = ActiveDocument.Paragraphs(1).Format.Duplicate 
Selection.Paragraphs.Format = paraFormat
```


## See also


[Paragraph Object](Word.Paragraph.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]