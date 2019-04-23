---
title: Paragraphs.TabHangingIndent method (Word)
keywords: vbawd10.chm156762416
f1_keywords:
- vbawd10.chm156762416
ms.prod: word
api_name:
- Word.Paragraphs.TabHangingIndent
ms.assetid: 6b99b0d8-15f9-1b44-3b97-f0f46e2757c1
ms.date: 06/08/2017
localization_priority: Normal
---


# Paragraphs.TabHangingIndent method (Word)

Sets a hanging indent to a specified number of tab stops.


## Syntax

_expression_. `TabHangingIndent`( `_Count_` )

_expression_ Required. A variable that represents a '[Paragraphs](Word.paragraphs.md)' collection.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Count_|Required| **Integer**|The number of tab stops to indent (if positive) or the number of tab stops to remove from the indent (if negative).|

## Remarks

You can also use this method to remove tab stops from a hanging indent if the value of Count is a negative number.


## Example

This example sets a hanging indent all paragraphs in the active document.


```vb
ActiveDocument.Paragraphs.TabHangingIndent(2)
```

This example moves the hanging indent back one tab stop for all paragraphs in the active document.




```vb
ActiveDocument.Paragraphs.TabHangingIndent(-1)
```


## See also


[Paragraphs Collection Object](Word.paragraphs.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]