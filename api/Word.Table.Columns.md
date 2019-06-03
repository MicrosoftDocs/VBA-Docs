---
title: Table.Columns property (Word)
keywords: vbawd10.chm156303460
f1_keywords:
- vbawd10.chm156303460
ms.prod: word
api_name:
- Word.Table.Columns
ms.assetid: 6f4c70ef-032d-7f05-1b21-c5c86af804bd
ms.date: 06/08/2017
localization_priority: Normal
---


# Table.Columns property (Word)

Returns a  **[Columns](Word.columns.md)** collection that represents all the table columns in the table. Read-only.


## Syntax

_expression_.**Columns**

_expression_ A variable that represents a '[Table](Word.Table.md)' object.


## Remarks

For information about returning a single member of a collection, see [Returning an object from a collection](../word/Concepts/Miscellaneous/returning-an-object-from-a-collection-word.md).


## Example

This example displays the number of columns in the first table in the active document.


```vb
If ActiveDocument.Tables.Count >= 1 Then 
 MsgBox ActiveDocument.Tables(1).Columns.Count 
End If
```


## See also


[Table Object](Word.Table.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]