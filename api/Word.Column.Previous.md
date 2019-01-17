---
title: Column.Previous property (Word)
keywords: vbawd10.chm156172392
f1_keywords:
- vbawd10.chm156172392
ms.prod: word
api_name:
- Word.Column.Previous
ms.assetid: 633b0d86-5591-5dcf-f2f3-f414c193b4cd
ms.date: 06/08/2017
localization_priority: Normal
---


# Column.Previous property (Word)

Returns the previous column in a collection of table columns. Read-only.


## Syntax

 _expression_. `Previous`

 _expression_ A variable that represents a '[Column](Word.Column.md)' object.


## Example

If the selection is in a table, this example selects the contents of the previous table column.


```vb
If Selection.Information(wdWithInTable) = True Then 
 Selection.Columns(1).Previous.Select 
End If
```


## See also


[Column Object](Word.Column.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]