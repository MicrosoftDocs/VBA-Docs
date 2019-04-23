---
title: Paragraphs.BaseLineAlignment property (Word)
keywords: vbawd10.chm156762235
f1_keywords:
- vbawd10.chm156762235
ms.prod: word
api_name:
- Word.Paragraphs.BaseLineAlignment
ms.assetid: 023055e7-62a0-475c-2f26-962d1c0f207c
ms.date: 06/08/2017
localization_priority: Normal
---


# Paragraphs.BaseLineAlignment property (Word)

Returns or sets a  **WdBaselineAlignment** constant that represents the vertical position of fonts on a line. Read/write.


## Syntax

_expression_. `BaseLineAlignment`

_expression_ Required. A variable that represents a '[Paragraphs](Word.paragraphs.md)' collection.


## Example

This example sets Microsoft Word to automatically adjust the baseline font alignment in the active document.


```vb
ActiveDocument.BaseLineAlignment = wdBaselineAlignAuto
```


## See also


[Paragraphs Collection Object](Word.paragraphs.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]