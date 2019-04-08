---
title: Footnotes.ResetSeparator method (Word)
keywords: vbawd10.chm155320327
f1_keywords:
- vbawd10.chm155320327
ms.prod: word
api_name:
- Word.Footnotes.ResetSeparator
ms.assetid: 252633ab-a9a1-6dbe-7821-5c7969175996
ms.date: 06/08/2017
localization_priority: Normal
---


# Footnotes.ResetSeparator method (Word)

Resets the footnote separator to the default separator.


## Syntax

_expression_. `ResetSeparator`

_expression_ Required. A variable that represents a '[Footnotes](Word.footnotes.md)' collection.


## Remarks

The default separator is a short horizontal line that separates document text from notes.


## Example

This example resets the footnote separator to the default separator line.


```vb
ActiveDocument.Footnotes.ResetSeparator
```


## See also


[Footnotes Collection Object](Word.footnotes.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]