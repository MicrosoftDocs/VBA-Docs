---
title: ParagraphFormat.Space15 method (Word)
keywords: vbawd10.chm156434746
f1_keywords:
- vbawd10.chm156434746
ms.prod: word
api_name:
- Word.ParagraphFormat.Space15
ms.assetid: 6621d8e8-c207-0862-ddd4-33cb5bcd9cbc
ms.date: 06/08/2017
localization_priority: Normal
---


# ParagraphFormat.Space15 method (Word)

Formats the specified paragraphs with 1.5-line spacing.


## Syntax

_expression_. `Space15`

_expression_ Required. A variable that represents a '[ParagraphFormat](Word.ParagraphFormat.md)' object.


## Remarks

The exact spacing is determined by adding 6 points to the font size of the largest character in each paragraph.

You can also use the  **[LineSpacingRule](Word.ParagraphFormat.LineSpacingRule.md)** property to set the spacing of paragraphs. The following two statements are equivalent:




```vb
Selection.ParagraphFormat.Space15 
Selection.ParagraphFormat.LineSpacingRule = wdLineSpace1pt5
```


## Example

This example changes the first paragraph in the active document to 1.5-line spacing.


```vb
Selection.ParagraphFormat.Space15
```


## See also


[ParagraphFormat Object](Word.ParagraphFormat.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]