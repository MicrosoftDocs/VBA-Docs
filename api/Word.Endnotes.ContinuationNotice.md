---
title: Endnotes.ContinuationNotice property (Word)
keywords: vbawd10.chm155254890
f1_keywords:
- vbawd10.chm155254890
ms.prod: word
api_name:
- Word.Endnotes.ContinuationNotice
ms.assetid: 3d2007df-756e-17f9-ce7c-269fa633503b
ms.date: 06/08/2017
localization_priority: Normal
---


# Endnotes.ContinuationNotice property (Word)

Returns a  **Range** object that represents the endnote continuation notice. Read-only.


## Syntax

_expression_. `ContinuationNotice`

_expression_ A variable that represents an '[Endnotes](Word.endnotes.md)' collection.


## Example

This example replaces the current footnote continuation notice with the text "Continued...".


```vb
With ActiveDocument.Footnotes.ContinuationNotice 
 .Delete 
 .InsertBefore "Continued..." 
End With
```


## See also


[Endnotes Collection Object](Word.endnotes.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]