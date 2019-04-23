---
title: Paragraph.OpenOrCloseUp method (Word)
keywords: vbawd10.chm156696879
f1_keywords:
- vbawd10.chm156696879
ms.prod: word
api_name:
- Word.Paragraph.OpenOrCloseUp
ms.assetid: ab5a657f-9a8f-a191-76ac-f16aaa2758ee
ms.date: 06/08/2017
localization_priority: Normal
---


# Paragraph.OpenOrCloseUp method (Word)

Toggles the spacing before a paragraph.


## Syntax

_expression_. `OpenOrCloseUp`

_expression_ Required. A variable that represents a '[Paragraph](Word.Paragraph.md)' object.


## Remarks

If spacing before the specified paragraphs is 0 (zero), this method sets spacing to 12 points. If spacing before the paragraphs is greater than 0 (zero), this method sets spacing to 0 (zero).


## Example

This example toggles the formatting of the first paragraph in the active document to either add 12 points of space before the paragraph or leave no space before it.


```vb
ActiveDocument.Paragraphs(1).OpenOrCloseUp
```


## See also


[Paragraph Object](Word.Paragraph.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]