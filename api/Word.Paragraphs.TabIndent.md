---
title: Paragraphs.TabIndent method (Word)
keywords: vbawd10.chm156762418
f1_keywords:
- vbawd10.chm156762418
ms.prod: word
api_name:
- Word.Paragraphs.TabIndent
ms.assetid: 37a7ea00-c9c5-c3a4-c01a-020f5cfd0ad7
ms.date: 06/08/2017
localization_priority: Normal
---


# Paragraphs.TabIndent method (Word)

Sets the left indent for the specified paragraphs to a specified number of tab stops.


## Syntax

_expression_. `TabIndent`( `_Count_` )

_expression_ Required. A variable that represents a '[Paragraphs](Word.paragraphs.md)' collection.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Count_|Required| **Integer**|The number of tab stops to indent (if positive) or the number of tab stops to remove from the indent (if negative).|

## Remarks

You can also use this method to remove the indent if the value of Count is a negative number.


## Example

This example indents all paragraphs in the active document to the second tab stop.


```vb
ActiveDocument.Paragraphs.TabIndent(2)
```

This example moves the indent of all paragraphs in the active document back one tab stop.




```vb
ActiveDocument.Paragraphs.TabIndent(-1)
```


## See also


[Paragraphs Collection Object](Word.paragraphs.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]