---
title: Range.TwoLinesInOne property (Word)
keywords: vbawd10.chm157155594
f1_keywords:
- vbawd10.chm157155594
ms.prod: word
api_name:
- Word.Range.TwoLinesInOne
ms.assetid: 08e91e95-4826-7df9-22a9-3c7b9c25042d
ms.date: 06/08/2017
localization_priority: Normal
---


# Range.TwoLinesInOne property (Word)

Returns or sets whether Microsoft Word sets two lines of text in one and specifies the characters that enclose the text, if any. Read/write  **WdTwoLinesInOneType**.


## Syntax

_expression_. `TwoLinesInOne`

_expression_ Required. A variable that represents a **[Range](Word.Range.md)** object.


## Remarks

Setting the **TwoLinesInOne** property to **wdTwoLinesInOneNoBrackets** sets two lines of text in one without enclosing the text in any characters. Setting the **TwoLinesInOne** property to **wdTwoLinesInOneNone** restores a line of combined text to two separate lines.


## Example

This example formats the current selection as two lines of text in one, enclosed in parentheses.


```vb
Selection.Range.TwoLinesInOne = _ 
 wdTwoLinesInOneParentheses
```


## See also


[Range Object](Word.Range.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]