---
title: Columns.Add method (Word)
keywords: vbawd10.chm155910149
f1_keywords:
- vbawd10.chm155910149
ms.prod: word
api_name:
- Word.Columns.Add
ms.assetid: b93aa859-e0f1-b8b1-a9d7-766f7f1f528c
ms.date: 06/08/2017
localization_priority: Normal
---


# Columns.Add method (Word)

Returns a  **Column** object that represents a column added to a table.


## Syntax

_expression_.**Add** (_BeforeColumn_)

_expression_ Required. A variable that represents a '[Columns](Word.columns.md)' collection.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _BeforeColumn_|Optional| **Variant**|A **Column** object that represents the column that will appear immediately to the right of the new column.|

## Return value

Column


## Example

This example creates a table with two columns and two rows in the active document and then adds another column before the first column. The width of the new column is set at 1.5 inches.


```vb
Sub AddATable() 
 Dim myTable As Table 
 Dim newCol As Column 
 
 Set myTable = ActiveDocument.Tables.Add(Selection.Range, 2, 2) 
 Set newCol = myTable.Columns.Add(BeforeColumn:=myTable.Columns(1)) 
 newCol.SetWidth ColumnWidth:=InchesToPoints(1.5), _ 
 RulerStyle:=wdAdjustNone 
End Sub
```


## See also


[Columns Collection Object](Word.columns.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]