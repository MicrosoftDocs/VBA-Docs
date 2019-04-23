---
title: Paragraphs.OpenOrCloseUp method (Word)
keywords: vbawd10.chm156762415
f1_keywords:
- vbawd10.chm156762415
ms.prod: word
api_name:
- Word.Paragraphs.OpenOrCloseUp
ms.assetid: b8531067-8c4a-e3aa-2561-aae4c20d7abf
ms.date: 06/08/2017
localization_priority: Normal
---


# Paragraphs.OpenOrCloseUp method (Word)

Toggles spacing before paragraphs.


## Syntax

_expression_. `OpenOrCloseUp`

_expression_ Required. A variable that represents a '[Paragraphs](Word.paragraphs.md)' collection.


## Remarks

If spacing before the specified paragraphs is 0 (zero), this method sets spacing to 12 points. If spacing before the paragraphs is greater than 0 (zero), this method sets spacing to 0 (zero).


## Example

This example toggles the formatting of the first paragraph in the active document to either add 12 points of space before the paragraph or leave no space before it.


```vb
ActiveDocument.Paragraphs(1).OpenOrCloseUp
```


## See also


[Paragraphs Collection Object](Word.paragraphs.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]