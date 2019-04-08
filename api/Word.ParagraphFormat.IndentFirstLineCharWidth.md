---
title: ParagraphFormat.IndentFirstLineCharWidth method (Word)
keywords: vbawd10.chm156434754
f1_keywords:
- vbawd10.chm156434754
ms.prod: word
api_name:
- Word.ParagraphFormat.IndentFirstLineCharWidth
ms.assetid: 9531e607-4287-d4a3-de85-315e806d9b51
ms.date: 06/08/2017
localization_priority: Normal
---


# ParagraphFormat.IndentFirstLineCharWidth method (Word)

Indents the first line of one or more paragraphs by a specified number of characters.


## Syntax

_expression_. `IndentFirstLineCharWidth`( `_Count_` )

_expression_ Required. A variable that represents a '[ParagraphFormat](Word.ParagraphFormat.md)' object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Count_|Required| **Integer**|The number of characters by which the first line of each specified paragraph is to be indented.|

## Example

This example indents the first line of the first paragraph in the active document by 10 characters.


```vb
Selection.ParagraphFormat.IndentFirstLineCharWidth 10
```


## See also


[ParagraphFormat Object](Word.ParagraphFormat.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]