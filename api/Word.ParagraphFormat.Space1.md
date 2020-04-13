---
title: ParagraphFormat.Space1 method (Word)
keywords: vbawd10.chm156434745
f1_keywords:
- vbawd10.chm156434745
ms.prod: word
api_name:
- Word.ParagraphFormat.Space1
ms.assetid: 57cc0cea-e50d-affd-1564-30f9240f197b
ms.date: 06/08/2017
localization_priority: Normal
---


# ParagraphFormat.Space1 method (Word)

Single-spaces the specified paragraphs.


## Syntax

_expression_. `Space1`

_expression_ Required. A variable that represents a '[ParagraphFormat](Word.ParagraphFormat.md)' object.


## Remarks

The exact spacing is determined by the font size of the largest characters in each paragraph.

You can also use the **[LineSpacingRule](Word.ParagraphFormat.LineSpacingRule.md)** property to set the spacing of paragraphs. The following two statements are equivalent:




```vb
Selection.ParagraphFormat.Space1 
Selection.ParagraphFormat.LineSpacingRule = wdLineSpaceSingle
```


## Example

This example changes the first paragraph in the active document to single spacing.


```vb
Selection.ParagraphFormat.Space1
```


## See also


[ParagraphFormat Object](Word.ParagraphFormat.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]