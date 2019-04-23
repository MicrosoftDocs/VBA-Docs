---
title: Endnotes.ResetSeparator method (Word)
keywords: vbawd10.chm155254791
f1_keywords:
- vbawd10.chm155254791
ms.prod: word
api_name:
- Word.Endnotes.ResetSeparator
ms.assetid: 9d525a4b-d3ed-5a31-9c07-1c19129cd171
ms.date: 06/08/2017
localization_priority: Normal
---


# Endnotes.ResetSeparator method (Word)

Resets the endnote separator to the default separator.


## Syntax

_expression_. `ResetSeparator`

_expression_ Required. A variable that represents an '[Endnotes](Word.endnotes.md)' collection.


## Remarks

 The default separator is a short horizontal line that separates document text from notes.


## Example

This example resets the endnote separator for the notes in the document where the selection is located.


```vb
Selection.Endnotes.ResetSeparator
```


## See also


[Endnotes Collection Object](Word.endnotes.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]