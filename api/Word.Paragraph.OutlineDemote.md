---
title: Paragraph.OutlineDemote method (Word)
keywords: vbawd10.chm156696903
f1_keywords:
- vbawd10.chm156696903
ms.prod: word
api_name:
- Word.Paragraph.OutlineDemote
ms.assetid: 02e65a97-6334-5205-b69e-a38f7aaeb8fd
ms.date: 06/08/2017
localization_priority: Normal
---


# Paragraph.OutlineDemote method (Word)

Applies the next heading level style (Heading 1 through Heading 8) to the specified paragraph or paragraphs.


## Syntax

_expression_. `OutlineDemote`

_expression_ Required. A variable that represents a '[Paragraph](Word.Paragraph.md)' object.


## Remarks

If a paragraph is formatted with the Heading 2 style, this method demotes the paragraph by changing the style to Heading 3.


## Example

This example demotes the first paragraph in the selection.


```vb
Selection.Paragraphs(1).OutlineDemote
```

This example demotes the third paragraph in the active document.




```vb
ActiveDocument.Paragraphs(3).OutlineDemote
```


## See also


[Paragraph Object](Word.Paragraph.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]