---
title: Footnotes.ContinuationSeparator property (Word)
keywords: vbawd10.chm155320425
f1_keywords:
- vbawd10.chm155320425
ms.prod: word
api_name:
- Word.Footnotes.ContinuationSeparator
ms.assetid: 5bcb180c-58a0-28e8-3712-7a1ee0e731b9
ms.date: 06/08/2017
localization_priority: Normal
---


# Footnotes.ContinuationSeparator property (Word)

Returns a  **Range** object that represents the footnote continuation separator. Read-only.


## Syntax

 _expression_. `ContinuationSeparator`

 _expression_ A variable that represents a '[Footnotes](Word.footnotes.md)' collection.


## Example

This example replaces the current endnote continuation separator with a series of underscore characters.


```vb
With ActiveDocument.Footnotes.ContinuationSeparator 
 .Delete 
 .InsertBefore "____" 
End With
```


## See also


[Footnotes Collection Object](Word.footnotes.md)

