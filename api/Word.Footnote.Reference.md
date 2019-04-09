---
title: Footnote.Reference property (Word)
keywords: vbawd10.chm155123717
f1_keywords:
- vbawd10.chm155123717
ms.prod: word
api_name:
- Word.Footnote.Reference
ms.assetid: c13dfad2-a103-8d91-0e55-86022a7857cd
ms.date: 06/08/2017
localization_priority: Normal
---


# Footnote.Reference property (Word)

Returns a  **[Range](Word.Range.md)** object that represents a footnote reference mark.


## Syntax

_expression_. `Reference`

_expression_ Required. A variable that represents a '[Footnote](Word.Footnote.md)' object.


## Example

This example sets  _myRange_ to the first footnote reference mark in the active document and then copies the reference mark.


```vb
Set myRange = ActiveDocument.Footnotes(1).Reference 
myRange.Copy
```


## See also


[Footnote Object](Word.Footnote.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]