---
title: Paragraphs.CloseUp method (Word)
keywords: vbawd10.chm156762413
f1_keywords:
- vbawd10.chm156762413
ms.prod: word
api_name:
- Word.Paragraphs.CloseUp
ms.assetid: 0fa0afb7-fbdf-ab26-1b49-312f526d69c6
ms.date: 06/08/2017
localization_priority: Normal
---


# Paragraphs.CloseUp method (Word)

Removes any spacing before the specified paragraphs.


## Syntax

_expression_. `CloseUp`

_expression_ Required. A variable that represents a '[Paragraphs](Word.paragraphs.md)' collection.


## Remarks

The following two statements are equivalent:


```vb
ActiveDocument.Paragraphs.CloseUp 
ActiveDocument.Paragraphs.SpaceBefore = 0
```


## Example

This example removes any space before the first paragraph in the selection.


```vb
Selection.Paragraphs.CloseUp
```


## See also


[Paragraphs Collection Object](Word.paragraphs.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]