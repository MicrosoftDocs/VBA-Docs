---
title: Endnotes.ContinuationSeparator property (Word)
keywords: vbawd10.chm155254889
f1_keywords:
- vbawd10.chm155254889
ms.prod: word
api_name:
- Word.Endnotes.ContinuationSeparator
ms.assetid: 4f62aa74-5c9e-6f95-ddc5-ff02c9a00bcf
ms.date: 06/08/2017
localization_priority: Normal
---


# Endnotes.ContinuationSeparator property (Word)

Returns a  **Range** object that represents the endnote continuation separator. Read-only.


## Syntax

_expression_. `ContinuationSeparator`

_expression_ A variable that represents an '[Endnotes](Word.endnotes.md)' collection.


## Example

This example replaces the current endnote continuation separator with a series of underscore characters.


```vb
With ActiveDocument.Endnotes.ContinuationSeparator 
 .Delete 
 .InsertBefore "____" 
End With
```


## See also


[Endnotes Collection Object](Word.endnotes.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]