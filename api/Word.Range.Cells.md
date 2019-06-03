---
title: Range.Cells property (Word)
keywords: vbawd10.chm157155385
f1_keywords:
- vbawd10.chm157155385
ms.prod: word
api_name:
- Word.Range.Cells
ms.assetid: aa081698-53d0-2234-5ec3-6e9a4091caef
ms.date: 06/08/2017
localization_priority: Normal
---


# Range.Cells property (Word)

Returns a  **[Cells](Word.cells.md)** collection that represents the table cells in a range. Read-only.


## Syntax

_expression_.**Cells**

_expression_ A variable that represents a **[Range](Word.Range.md)** object.


## Remarks

For information about returning a single member of a collection, see [Returning an object from a collection](../word/Concepts/Miscellaneous/returning-an-object-from-a-collection-word.md).


## Example

This example creates a 3x3 table and assigns a sequential cell number to each cell in the table.


```vb
Set newDoc = Documents.Add 
Set myTable = newDoc.Tables.Add(Selection.Range, 3, 3) 
i = 1 
For Each c In myTable.Range.Cells 
 c.Range.InsertAfter "Cell " & i 
 i = i + 1 
Next c
```


## See also


[Range Object](Word.Range.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]