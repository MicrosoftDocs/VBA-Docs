---
title: Column.IsLast property (Word)
keywords: vbawd10.chm156172293
f1_keywords:
- vbawd10.chm156172293
ms.prod: word
api_name:
- Word.Column.IsLast
ms.assetid: 9f5e51fe-4bb7-a179-4dde-373f7798f200
ms.date: 06/08/2017
localization_priority: Normal
---


# Column.IsLast property (Word)

 **True** if the specified column or row is the last one in the table. Read-only **Boolean**.


## Syntax

_expression_. `IsLast`

_expression_ Required. A variable that represents a '[Column](Word.Column.md)' object.


## Example

This example determines whether the first column in the selection is the last column in the table.


```vb
If Selection.Information(wdWithInTable) = True Then 
 MsgBox Selection.Columns(1).IsLast 
End If
```


## See also


[Column Object](Word.Column.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]