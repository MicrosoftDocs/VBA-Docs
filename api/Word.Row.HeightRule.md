---
title: Row.HeightRule property (Word)
keywords: vbawd10.chm156237832
f1_keywords:
- vbawd10.chm156237832
ms.prod: word
api_name:
- Word.Row.HeightRule
ms.assetid: 7dad51e9-e819-6c7b-a562-7e3b7ca58f3c
ms.date: 06/08/2017
localization_priority: Normal
---


# Row.HeightRule property (Word)

Returns or sets the rule for determining the height of the specified cells or rows. Read/write  **WdRowHeightRule**.


## Syntax

_expression_. `HeightRule`

_expression_ Required. A variable that represents a '[Row](Word.Row.md)' object.


## Example

This example creates a 3x3 table in a new document and then sets a minimum row height of 24 points for the second row.


```vb
Set newDoc = Documents.Add 
Set myTable = newDoc.Tables.Add(Range:=Selection.Range, _ 
 NumRows:=3, NumColumns:=3) 
With myTable.Rows(2) 
 .Height = 24 
 .HeightRule = wdRowHeightAtLeast 
End With
```


## See also


[Row Object](Word.Row.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]