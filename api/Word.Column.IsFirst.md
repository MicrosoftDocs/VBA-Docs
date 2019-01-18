---
title: Column.IsFirst property (Word)
keywords: vbawd10.chm156172292
f1_keywords:
- vbawd10.chm156172292
ms.prod: word
api_name:
- Word.Column.IsFirst
ms.assetid: 415048d5-b7a8-67e5-674b-19ca8ba93d8a
ms.date: 06/08/2017
localization_priority: Normal
---


# Column.IsFirst property (Word)

 **True** if the specified column or row is the first one in the table. Read-only **Boolean**.


## Syntax

 _expression_. `IsFirst`

 _expression_ Required. A variable that represents a '[Column](Word.Column.md)' object.


## Example

This example indicates whether the first column in the selection is the first column in the table.


```vb
MsgBox Selection.Columns(1).IsFirst
```


## See also


[Column Object](Word.Column.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]