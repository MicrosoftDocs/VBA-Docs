---
title: Table.Cell method (Word)
keywords: vbawd10.chm156303377
f1_keywords:
- vbawd10.chm156303377
ms.prod: word
api_name:
- Word.Table.Cell
ms.assetid: 7dd91771-c72b-eefb-2492-1998c0d194bb
ms.date: 06/08/2017
localization_priority: Normal
---


# Table.Cell method (Word)

Returns a  **Cell** object that represents a cell in a table.


## Syntax

_expression_. `Cell`( `_Row_` , `_Column_` )

_expression_ Required. A variable that represents a '[Table](Word.Table.md)' object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Row_|Required| **Long**|The number of the row in the table to return. Can be an integer between 1 and the number of rows in the table.|
| _Column_|Required| **Long**|The number of the cell in the table to return. Can be an integer between 1 and the number of columns in the table.|

## Return value

Cell


## Example

This example creates a 3x3 table in a new document and inserts text into the first and last cells in the table.


```vb
Dim docNew As Document 
Dim tableNew As Table 
 
Set docNew = Documents.Add 
Set tableNew = docNew.Tables.Add(Selection.Range, 3, 3) 
 
With tableNew 
 .Cell(1,1).Range.InsertAfter "First cell" 
 .Cell(tableNew.Rows.Count, _ 
 tableNew.Columns.Count).Range.InsertAfter "Last Cell" 
End With
```

This example deletes the first cell from the first table in the active document.




```vb
If ActiveDocument.Tables.Count >= 1 Then 
 ActiveDocument.Tables(1).Cell(1, 1).Delete 
End If
```


## See also


[Table Object](Word.Table.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
