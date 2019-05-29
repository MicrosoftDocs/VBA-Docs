---
title: Options.DefaultBorderColor property (Word)
keywords: vbawd10.chm162988376
f1_keywords:
- vbawd10.chm162988376
ms.prod: word
api_name:
- Word.Options.DefaultBorderColor
ms.assetid: 382f9780-d10d-925b-206d-d7c624b6b744
ms.date: 06/08/2017
localization_priority: Normal
---


# Options.DefaultBorderColor property (Word)

Returns or sets the default 24-bit color to use for new  **[Border](Word.Border.md)** objects. Read/write.


## Syntax

_expression_. `DefaultBorderColor`

_expression_ Required. A variable that represents an **[Options](Word.Options.md)** object.


## Remarks

This property can be any valid  **WdColor** constant or a value returned by Visual Basic's **RGB** function.


## Example

This example sets the default color for new borders to teal.


```vb
Options.DefaultBorderColor = wdColorTeal
```


## See also


[Options Object](Word.Options.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]