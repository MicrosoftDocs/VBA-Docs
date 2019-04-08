---
title: ParagraphFormat.TabIndent method (Word)
keywords: vbawd10.chm156434738
f1_keywords:
- vbawd10.chm156434738
ms.prod: word
api_name:
- Word.ParagraphFormat.TabIndent
ms.assetid: db62f9c2-e205-4f57-5baf-2c06bbd30644
ms.date: 06/08/2017
localization_priority: Normal
---


# ParagraphFormat.TabIndent method (Word)

Sets the left indent for the specified paragraphs to a specified number of tab stops.


## Syntax

_expression_. `TabIndent`( `_Count_` )

_expression_ Required. A variable that represents a '[ParagraphFormat](Word.ParagraphFormat.md)' object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Count_|Required| **Integer**|The number of tab stops to indent (if positive) or the number of tab stops to remove from the indent (if negative).|

## Remarks

You can also use this method to remove an indent if the value of Count is a negative number.


## Example

This example indents the selected paragraphs to the second tab stop.


```vb
Selection.ParagraphFormat.TabIndent(2)
```

This example moves the indent of the selected paragraphs back one tab stop.




```vb
Selection.ParagraphFormat.TabIndent(-1)
```


## See also


[ParagraphFormat Object](Word.ParagraphFormat.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]