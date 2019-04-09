---
title: Footnotes.ResetContinuationSeparator method (Word)
keywords: vbawd10.chm155320328
f1_keywords:
- vbawd10.chm155320328
ms.prod: word
api_name:
- Word.Footnotes.ResetContinuationSeparator
ms.assetid: edb1dae6-3e62-b625-0982-64dec3b654c9
ms.date: 06/08/2017
localization_priority: Normal
---


# Footnotes.ResetContinuationSeparator method (Word)

Resets the footnote or endnote continuation separator to the default separator.


## Syntax

_expression_. `ResetContinuationSeparator`

_expression_ Required. A variable that represents a '[Footnotes](Word.footnotes.md)' collection.


## Remarks

The default separator is a long horizontal line that separates document text from notes continued from the previous page.


## Example

This example resets the footnote continuation separator to the default separator line.


```vb
ActiveDocument.Footnotes.ResetContinuationSeparator
```


## See also


[Footnotes Collection Object](Word.footnotes.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]