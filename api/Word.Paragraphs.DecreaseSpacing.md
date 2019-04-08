---
title: Paragraphs.DecreaseSpacing method (Word)
keywords: vbawd10.chm156762448
f1_keywords:
- vbawd10.chm156762448
ms.prod: word
api_name:
- Word.Paragraphs.DecreaseSpacing
ms.assetid: 9d1dfab7-87a0-21c0-f023-0b1368aa9773
ms.date: 06/08/2017
localization_priority: Normal
---


# Paragraphs.DecreaseSpacing method (Word)

Decreases the spacing before and after paragraphs in six-point increments.


## Syntax

_expression_. `DecreaseSpacing`

_expression_ Required. A variable that represents a '[Paragraphs](Word.paragraphs.md)' collection.


## Example

This example decreases the before and after spacing of a paragraph or selection of paragraphs by six points each time the procedure is run. If the before and after spacing are both zero, the procedure will do nothing.


```vb
Sub DecreaseParaSpacing() 
 Selection.Paragraphs.DecreaseSpacing 
End Sub
```


## See also


[Paragraphs Collection Object](Word.paragraphs.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]