---
title: Paragraphs.IndentFirstLineCharWidth method (Word)
keywords: vbawd10.chm156762434
f1_keywords:
- vbawd10.chm156762434
ms.prod: word
api_name:
- Word.Paragraphs.IndentFirstLineCharWidth
ms.assetid: d0fc2250-8e3a-8a35-7d15-2bd9cc3653db
ms.date: 06/08/2017
localization_priority: Normal
---


# Paragraphs.IndentFirstLineCharWidth method (Word)

Indents the first line of one or more paragraphs by a specified number of characters.


## Syntax

_expression_. `IndentFirstLineCharWidth`( `_Count_` )

_expression_ Required. A variable that represents a '[Paragraphs](Word.paragraphs.md)' collection.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Count_|Required| **Integer**|The number of characters by which the first line of each specified paragraph is to be indented.|

## Example

This example indents the first line of all paragraphs in the active document by 10 characters.


```vb
With ActiveDocument.Paragraphs 
 .IndentFirstLineCharWidth 10 
End With
```


## See also


[Paragraphs Collection Object](Word.paragraphs.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]