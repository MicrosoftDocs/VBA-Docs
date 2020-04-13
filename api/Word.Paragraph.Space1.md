---
title: Paragraph.Space1 method (Word)
keywords: vbawd10.chm156696889
f1_keywords:
- vbawd10.chm156696889
ms.prod: word
api_name:
- Word.Paragraph.Space1
ms.assetid: e9eaab54-d910-f6d4-8e7b-c47a7395d00b
ms.date: 06/08/2017
localization_priority: Normal
---


# Paragraph.Space1 method (Word)

Single-spaces the specified paragraphs.


## Syntax

_expression_. `Space1`

_expression_ Required. A variable that represents a '[Paragraph](Word.Paragraph.md)' object.


## Remarks

The exact spacing is determined by the font size of the largest characters in each paragraph.

You can also use the **[LineSpacingRule](Word.Paragraph.LineSpacingRule.md)** property to set the line spacing for a paragraph. The following two statements are equivalent:




```vb
ActiveDocument.Paragraphs(1).Space1 
ActiveDocument.Paragraphs(1).LineSpacingRule = wdLineSpaceSingle
```


## Example

This example changes the first paragraph in the active document to single spacing.


```vb
ActiveDocument.Paragraphs(1).Space1
```


## See also


[Paragraph Object](Word.Paragraph.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]