---
title: ParagraphFormat.OpenUp method (Word)
keywords: vbawd10.chm156434734
f1_keywords:
- vbawd10.chm156434734
ms.prod: word
api_name:
- Word.ParagraphFormat.OpenUp
ms.assetid: 1473b383-816f-087a-073a-5afc5f530c3a
ms.date: 06/08/2017
localization_priority: Normal
---


# ParagraphFormat.OpenUp method (Word)

Sets spacing before the specified paragraphs to 12 points.


## Syntax

_expression_. `OpenUp`

_expression_ Required. A variable that represents a '[ParagraphFormat](Word.ParagraphFormat.md)' object.


## Remarks

You can also use the **[SpaceBefore](Word.ParagraphFormat.SpaceBefore.md)** property to set the spacing of paragraphs. The following two statements are equivalent:


```vb
Selection.ParagraphFormat.OpenUp 
Selection.ParagraphFormat.SpaceBefore = 12
```


## Example

This example changes the formatting of the second paragraph in the active document to leave 12 points of space before the paragraph.


```vb
Selection.ParagraphFormat.OpenUp
```


## See also


[ParagraphFormat Object](Word.ParagraphFormat.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]