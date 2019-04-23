---
title: Paragraphs.OpenUp method (Word)
keywords: vbawd10.chm156762414
f1_keywords:
- vbawd10.chm156762414
ms.prod: word
api_name:
- Word.Paragraphs.OpenUp
ms.assetid: 0998519f-5fdc-3ac1-488f-03ff179be1c9
ms.date: 06/08/2017
localization_priority: Normal
---


# Paragraphs.OpenUp method (Word)

Sets spacing before the specified paragraphs to 12 points.


## Syntax

_expression_. `OpenUp`

_expression_ Required. A variable that represents a '[Paragraphs](Word.paragraphs.md)' collection.


## Remarks

You can also use the  **[SpaceBefore](Word.Paragraphs.SpaceBefore.md)** property to set the spacing before paragraphs. The following two statements are equivalent:


```vb
ActiveDocument.Paragraphs.OpenUp 
ActiveDocument.Paragraphs.SpaceBefore = 12
```


## Example

This example changes the formatting of the second paragraph in the active document to leave 12 points of space before the paragraph.


```vb
ActiveDocument.Paragraphs(2).OpenUp
```


## See also


[Paragraphs Collection Object](Word.paragraphs.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]