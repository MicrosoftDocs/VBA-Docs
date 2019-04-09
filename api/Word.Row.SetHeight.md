---
title: Row.SetHeight method (Word)
keywords: vbawd10.chm156238027
f1_keywords:
- vbawd10.chm156238027
ms.prod: word
api_name:
- Word.Row.SetHeight
ms.assetid: cbf4a6b3-b025-775e-d4c3-e5aa3c789522
ms.date: 06/08/2017
localization_priority: Normal
---


# Row.SetHeight method (Word)

Sets the height of a table row.


## Syntax

_expression_. `SetHeight`( `_RowHeight_` , `_HeightRule_` )

_expression_ Required. A variable that represents a '[Row](Word.Row.md)' object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _RowHeight_|Required| **Single**|The height of the row, in points.|
| _HeightRule_|Required| **WdRowHeightRule**|The rule for determining the height of the specified rows.|

## Example

This example creates a table and then sets a fixed row height of 0.5 inch (36 points) for the first row.


```vb
Set newDoc = Documents.Add 
Set aTable = _ 
 newDoc.Tables.Add(Range:=Selection.Range, NumRows:=3, _ 
 NumColumns:=3) 
aTable.Rows(1).SetHeight RowHeight:=InchesToPoints(0.5), _ 
 HeightRule:=wdRowHeightExactly
```


## See also


[Row Object](Word.Row.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]