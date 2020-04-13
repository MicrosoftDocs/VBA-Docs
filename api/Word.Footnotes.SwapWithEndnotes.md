---
title: Footnotes.SwapWithEndnotes method (Word)
keywords: vbawd10.chm155320326
f1_keywords:
- vbawd10.chm155320326
ms.prod: word
api_name:
- Word.Footnotes.SwapWithEndnotes
ms.assetid: ca92057d-e4f4-8931-83ad-73799fe830ea
ms.date: 06/08/2017
localization_priority: Normal
---


# Footnotes.SwapWithEndnotes method (Word)

Converts all footnotes in a document to endnotes and vice versa.To convert a range of footnotes to endnotes, use the **Convert** method.


## Syntax

_expression_. `SwapWithEndnotes`

_expression_ Required. A variable that represents a '[Footnotes](Word.footnotes.md)' collection.


## Example

This example converts the footnotes in the active document to endnotes and converts the endnotes to footnotes.


```vb
ActiveDocument.Footnotes.SwapWithEndnotes
```


## See also


[Footnotes Collection Object](Word.footnotes.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]