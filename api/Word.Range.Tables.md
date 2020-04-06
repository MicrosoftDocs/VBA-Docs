---
title: Range.Tables property (Word)
keywords: vbawd10.chm157155378
f1_keywords:
- vbawd10.chm157155378
ms.prod: word
api_name:
- Word.Range.Tables
ms.assetid: 1c6604be-233c-efb2-5d05-63fc5aa78481
ms.date: 06/08/2017
localization_priority: Normal
---


# Range.Tables property (Word)

Returns a  **Tables** collection that represents all the tables in the specified range. Read-only.


## Syntax

_expression_. `Tables`

_expression_ A variable that represents a **[Range](Word.Range.md)** object.


## Remarks

For information about returning a single member of a collection, see [Returning an object from a collection](../word/Concepts/Miscellaneous/returning-an-object-from-a-collection-word.md).


## Example

This example creates a 5x5 table in the active document and then applies a predefined format to it.


```vb
Selection.Collapse Direction:=wdCollapseStart 
Set myTable = ActiveDocument.Tables.Add(Range:=Selection.Range, _ 
NumRows:=5, NumColumns:=5) 
myTable.AutoFormat Format:=wdTableFormatClassic2
```

This example inserts numbers and text into the first column of the first table in the active document.




```vb
num = 90 
For Each acell In ActiveDocument.Tables(1).Columns(1).Cells 
 acell.Range.Text = num & " Sales" 
 num = num + 1 
Next acell
```


## See also


[Range Object](Word.Range.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]