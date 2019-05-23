---
title: Selection.SelectCurrentSpacing method (Word)
keywords: vbawd10.chm158663175
f1_keywords:
- vbawd10.chm158663175
ms.prod: word
api_name:
- Word.Selection.SelectCurrentSpacing
ms.assetid: 1a49caa6-d261-e9d7-9d64-c564c30a7e29
ms.date: 06/08/2017
localization_priority: Normal
---


# Selection.SelectCurrentSpacing method (Word)

Extends the selection forward until a paragraph with different line spacing is encountered.


## Syntax

_expression_. `SelectCurrentSpacing`

_expression_ Required. A variable that represents a **[Selection](Word.Selection.md)** object.


## Example

This example selects all consecutive paragraphs that have the same line spacing and changes the line spacing to single spacing.


```vb
With Selection 
 .SelectCurrentSpacing 
 .ParagraphFormat.Space1 
End With
```


## See also


[Selection Object](Word.Selection.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]