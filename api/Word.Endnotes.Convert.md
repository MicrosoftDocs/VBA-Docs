---
title: Endnotes.Convert method (Word)
keywords: vbawd10.chm155254789
f1_keywords:
- vbawd10.chm155254789
ms.prod: word
api_name:
- Word.Endnotes.Convert
ms.assetid: f351e0f2-ec4c-a9db-a119-1ebe3bb67319
ms.date: 06/08/2017
localization_priority: Normal
---


# Endnotes.Convert method (Word)

Converts endnotes to footnotes.


## Syntax

_expression_. `Convert`

_expression_ Required. A variable that represents an '[Endnotes](Word.endnotes.md)' object.


## Example

This example converts all endnotes in the active document to footnotes.


```vb
Set endDocEndnotes = ActiveDocument.Endnotes 
If endDocEndnotes.Count > 0 Then myEndnotes.Convert
```


## See also


[Endnotes Collection Object](Word.endnotes.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]