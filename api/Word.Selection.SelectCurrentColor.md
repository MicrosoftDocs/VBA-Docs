---
title: Selection.SelectCurrentColor method (Word)
keywords: vbawd10.chm158663178
f1_keywords:
- vbawd10.chm158663178
ms.prod: word
api_name:
- Word.Selection.SelectCurrentColor
ms.assetid: f7d23b80-7e1a-40a5-b292-820c3db500a6
ms.date: 06/08/2017
localization_priority: Normal
---


# Selection.SelectCurrentColor method (Word)

Extends the selection forward until text with a different color is encountered.


## Syntax

_expression_. `SelectCurrentColor`

_expression_ Required. A variable that represents a **[Selection](Word.Selection.md)** object.


## Example

This example extends the selection from the beginning of the document to the first character formatted with a different color and then displays the number of characters in the resulting selection.


```vb
Selection.HomeKey Unit:=wdStory, Extend:=wdMove 
Selection.SelectCurrentColor 
n = Len(Selection.Text) 
MsgBox "Contiguous characters with the same color: " & n
```


## See also


[Selection Object](Word.Selection.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]