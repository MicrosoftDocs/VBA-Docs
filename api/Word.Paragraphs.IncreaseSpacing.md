---
title: Paragraphs.IncreaseSpacing method (Word)
keywords: vbawd10.chm156762447
f1_keywords:
- vbawd10.chm156762447
ms.prod: word
api_name:
- Word.Paragraphs.IncreaseSpacing
ms.assetid: d0416601-5616-0e93-540f-f09e192b0c91
ms.date: 06/08/2017
localization_priority: Normal
---


# Paragraphs.IncreaseSpacing method (Word)

Increases the spacing before and after paragraphs in six-point increments.


## Syntax

_expression_. `IncreaseSpacing`

_expression_ Required. A variable that represents a '[Paragraphs](Word.paragraphs.md)' collection.


## Example

This example increases the before and after spacing of a paragraph or selection of paragraphs by six points each time the procedure is run.


```vb
Sub IncreaseParaSpacing() 
 Selection.Paragraphs.IncreaseSpacing 
End Sub
```


## See also


[Paragraphs Collection Object](Word.paragraphs.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]