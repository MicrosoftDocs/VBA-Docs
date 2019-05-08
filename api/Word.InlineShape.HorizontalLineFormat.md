---
title: InlineShape.HorizontalLineFormat property (Word)
keywords: vbawd10.chm162005111
f1_keywords:
- vbawd10.chm162005111
ms.prod: word
api_name:
- Word.InlineShape.HorizontalLineFormat
ms.assetid: 3e6f3887-d906-a761-d1ee-a4c4560c4888
ms.date: 06/08/2017
localization_priority: Normal
---


# InlineShape.HorizontalLineFormat property (Word)

Returns a  **[HorizontalLineFormat](Word.HorizontalLineFormat.md)** object that contains the horizontal line formatting for the specified **InlineShape** object. Read-only.


## Syntax

_expression_. `HorizontalLineFormat`

_expression_ A variable that represents a '[InlineShape](Word.InlineShape.md)' object.


## Example

This example sets the length of the specified horizontal line to 50% of the window width.


```vb
ActiveDocument.InlineShapes(1).HorizontalLineFormat _ 
 .PercentWidth = 50
```


## See also


[InlineShape Object](Word.InlineShape.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]