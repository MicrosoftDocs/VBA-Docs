---
title: Range.Columns property (Word)
keywords: vbawd10.chm157155630
f1_keywords:
- vbawd10.chm157155630
ms.prod: word
api_name:
- Word.Range.Columns
ms.assetid: 667b808a-e885-a7b7-0a68-5b2466ddd869
ms.date: 06/08/2017
localization_priority: Normal
---


# Range.Columns property (Word)

Returns a  **[Columns](Word.columns.md)** collection that represents all the table columns in the range. Read-only.


## Syntax

_expression_.**Columns**

_expression_ A variable that represents a **[Range](Word.Range.md)** object.


## Remarks

For information about returning a single member of a collection, see [Returning an object from a collection](../word/Concepts/Miscellaneous/returning-an-object-from-a-collection-word.md).


## Example

This example displays the number of columns in the first table in the active document.


```vb
If ActiveDocument.Tables.Count >= 1 Then 
 MsgBox ActiveDocument.Tables(1).Columns.Count 
End If
```

This example sets the width of the current column to 1 inch.




```vb
If Selection.Information(wdWithInTable) = True Then 
 Selection.Columns.SetWidth ColumnWidth:=InchesToPoints(1), _ 
 RulerStyle:=wdAdjustProportional 
End If
```


## See also


[Range Object](Word.Range.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]