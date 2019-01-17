---
title: Cell.Merge method (Word)
keywords: vbawd10.chm156106956
f1_keywords:
- vbawd10.chm156106956
ms.prod: word
api_name:
- Word.Cell.Merge
ms.assetid: 79d929bd-9578-e937-405f-8ad970ae883c
ms.date: 06/08/2017
localization_priority: Priority
---


# Cell.Merge method (Word)

Merges the specified table cell with another table cell. The result is a single table cell.


## Syntax

 _expression_. `Merge`( `_MergeTo_` )

 _expression_ Required. A variable that represents a '[Cell](Word.Cell.md)' object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _MergeTo_|Required| **Cell object**|The cell to be merged with.|

## Example

This example merges the first two cells in table one in the active document with one another and then removes the table borders.


```vb
If ActiveDocument.Tables.Count >= 1 Then 
 With ActiveDocument.Tables(1) 
 .Cell(Row:=1, Column:=1).Merge _ 
 MergeTo:=.Cell(Row:=1, Column:=2) 
 .Borders.Enable = False 
 End With 
End If
```


## See also


[Cell Object](Word.Cell.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]