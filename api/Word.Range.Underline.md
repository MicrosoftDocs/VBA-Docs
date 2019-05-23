---
title: Range.Underline property (Word)
keywords: vbawd10.chm157155467
f1_keywords:
- vbawd10.chm157155467
ms.prod: word
api_name:
- Word.Range.Underline
ms.assetid: 8221338d-3da6-b1ae-c424-87f762b61bd7
ms.date: 06/08/2017
localization_priority: Normal
---


# Range.Underline property (Word)

Returns or sets the type of underline applied to a range. Read/write  **WdUnderline**.


## Syntax

_expression_.**Underline**

_expression_ Required. A variable that represents a **[Range](Word.Range.md)** object.


## Example

This example applies a double underline to the fourth word in the active document.


```vb
ActiveDocument.Words(4).Underline = wdUnderlineDouble
```


## See also


[Range Object](Word.Range.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]