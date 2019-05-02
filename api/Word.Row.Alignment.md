---
title: Row.Alignment property (Word)
keywords: vbawd10.chm156237828
f1_keywords:
- vbawd10.chm156237828
ms.prod: word
api_name:
- Word.Row.Alignment
ms.assetid: 56214c5a-55d4-bcc9-857a-6591622bd264
ms.date: 06/08/2017
localization_priority: Normal
---


# Row.Alignment property (Word)

Returns or sets a  **WdRowAlignment** constant that represents the alignment for the specified rows. Read/write.


## Syntax

_expression_.**Alignment**

_expression_ Required. A variable that represents a '[Row](Word.Row.md)' object.


## Example

This example centers all the cells of the first row in the first table of the active document.


```vb
Sub CenterRows() 
 ActiveDocument.Tables(1).Rows(1) _ 
 .Alignment = wdAlignRowCenter 
End Sub
```


## See also


[Row Object](Word.Row.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]