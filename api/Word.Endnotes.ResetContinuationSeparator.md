---
title: Endnotes.ResetContinuationSeparator method (Word)
keywords: vbawd10.chm155254792
f1_keywords:
- vbawd10.chm155254792
ms.prod: word
api_name:
- Word.Endnotes.ResetContinuationSeparator
ms.assetid: 92de72c3-ab86-77e8-5047-928c145560cf
ms.date: 06/08/2017
localization_priority: Normal
---


# Endnotes.ResetContinuationSeparator method (Word)

Resets the endnote continuation separator to the default separator.


## Syntax

_expression_. `ResetContinuationSeparator`

_expression_ Required. A variable that represents an '[Endnotes](Word.endnotes.md)' collection.


## Remarks

The default separator is a long horizontal line that separates document text from notes continued from the previous page.


## Example

This example resets the endnote continuation separator for the first section in each open document.


```vb
Dim docLoop As Document 
 
For Each docLoop In Documents 
 docLoop.Sections(1).Range.Endnotes _ 
 .ResetContinuationSeparator 
Next docLoop
```


## See also


[Endnotes Collection Object](Word.endnotes.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]