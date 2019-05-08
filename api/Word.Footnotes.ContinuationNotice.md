---
title: Footnotes.ContinuationNotice property (Word)
keywords: vbawd10.chm155320426
f1_keywords:
- vbawd10.chm155320426
ms.prod: word
api_name:
- Word.Footnotes.ContinuationNotice
ms.assetid: 355a8bc1-3cf6-51e7-27f6-f3ff2b708fca
ms.date: 06/08/2017
localization_priority: Normal
---


# Footnotes.ContinuationNotice property (Word)

Returns a  **Range** object that represents the footnote continuation notice. Read-only.


## Syntax

_expression_. `ContinuationNotice`

_expression_ A variable that represents a '[Footnotes](Word.footnotes.md)' collection.


## Example

This example replaces the current footnote continuation notice with the text "Continued...".


```vb
With ActiveDocument.Footnotes.ContinuationNotice 
 .Delete 
 .InsertBefore "Continued..." 
End With
```


## See also


[Footnotes Collection Object](Word.footnotes.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]