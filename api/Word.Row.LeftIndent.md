---
title: Row.LeftIndent property (Word)
keywords: vbawd10.chm156237833
f1_keywords:
- vbawd10.chm156237833
ms.prod: word
api_name:
- Word.Row.LeftIndent
ms.assetid: 64dc0ca7-fd32-7dca-a09a-514af314c974
ms.date: 06/08/2017
localization_priority: Normal
---


# Row.LeftIndent property (Word)

Returns or sets a  **Single** that represents the left indent value (in points) for the specified table row. Read/write.


## Syntax

_expression_. `LeftIndent`

_expression_ A variable that represents a '[Row](Word.Row.md)' object.


## Example

This example sets the left indent for the first row in the first table in the active document.


```vb
ActiveDocument.Tables(1).Rows(1).LeftIndent = InchesToPoints(1)
```


## See also


[Row Object](Word.Row.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]