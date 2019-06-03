---
title: Document.Tables property (Word)
keywords: vbawd10.chm158007302
f1_keywords:
- vbawd10.chm158007302
ms.prod: word
api_name:
- Word.Document.Tables
ms.assetid: a0e09aff-af98-5d10-ba49-01ba6fcfa2d1
ms.date: 06/08/2017
localization_priority: Normal
---


# Document.Tables property (Word)

Returns a  **[Table](Word.Table.md)** collection that represents all the tables in the specified document. Read-only.


## Syntax

_expression_. `Tables`

_expression_ A variable that represents a **[Document](Word.Document.md)** object.


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


[Document Object](Word.Document.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]