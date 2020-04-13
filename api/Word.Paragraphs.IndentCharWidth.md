---
title: Paragraphs.IndentCharWidth method (Word)
keywords: vbawd10.chm156762432
f1_keywords:
- vbawd10.chm156762432
ms.prod: word
api_name:
- Word.Paragraphs.IndentCharWidth
ms.assetid: b463c523-8c2a-0609-db53-03238b4d232a
ms.date: 06/08/2017
localization_priority: Normal
---


# Paragraphs.IndentCharWidth method (Word)

Indents one or more paragraphs by a specified number of characters.


## Syntax

_expression_. `IndentCharWidth`( `_Count_` )

_expression_ Required. A variable that represents a '[Paragraphs](Word.paragraphs.md)' collection.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Count_|Required| **Integer**|The number of characters by which the specified paragraphs are to be indented.|

## Remarks

This method is equivalent to clicking the **Increase Indent** button on the **Formatting** toolbar.


## Example

This example indents all paragraphs in the active document by 10 characters.


```vb
With ActiveDocument.Paragraphs 
 .IndentCharWidth 10 
End With
```


## See also


[Paragraphs Collection Object](Word.paragraphs.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]