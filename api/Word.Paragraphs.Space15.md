---
title: Paragraphs.Space15 method (Word)
keywords: vbawd10.chm156762426
f1_keywords:
- vbawd10.chm156762426
ms.prod: word
api_name:
- Word.Paragraphs.Space15
ms.assetid: c48cb161-ba78-3fb6-bfb8-d13b6ec7e54d
ms.date: 06/08/2017
localization_priority: Normal
---


# Paragraphs.Space15 method (Word)

Formats the specified paragraphs with 1.5-line spacing.


## Syntax

_expression_. `Space15`

_expression_ Required. A variable that represents a '[Paragraphs](Word.paragraphs.md)' collection.


## Remarks

The exact spacing is determined by adding 6 points to the font size of the largest character in each paragraph.

You can also use the  **[LineSpacingRule](Word.Paragraphs.LineSpacingRule.md)** property to set paragraph spacing. The following two statements are equivalent:




```vb
ActiveDocument.Paragraphs.Space15 
ActiveDocument.Paragraphs.LineSpacingRule = wdLineSpace1pt5
```


## Example

This example changes all paragraphs in the active document to 1.5-line spacing.


```vb
ActiveDocument.Paragraphs.Space15
```


## See also


[Paragraphs Collection Object](Word.paragraphs.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]