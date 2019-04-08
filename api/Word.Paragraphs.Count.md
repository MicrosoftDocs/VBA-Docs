---
title: Paragraphs.Count property (Word)
keywords: vbawd10.chm156762114
f1_keywords:
- vbawd10.chm156762114
ms.prod: word
api_name:
- Word.Paragraphs.Count
ms.assetid: 8e2844f2-1a09-63d9-a981-e39a32a87d2f
ms.date: 06/08/2017
localization_priority: Normal
---


# Paragraphs.Count property (Word)

Returns a  **Long** that represents the number of paragraphs in the collection. Read-only.


## Syntax

_expression_.**Count**

_expression_ Required. A variable that represents a '[Paragraphs](Word.paragraphs.md)' collection.


## Example

This example displays the number of paragraphs in the active document.


```vb
MsgBox "The active document contains " & _ 
 ActiveDocument.Paragraphs.Count & " paragraphs."
```


## See also


[Paragraphs Collection Object](Word.paragraphs.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]