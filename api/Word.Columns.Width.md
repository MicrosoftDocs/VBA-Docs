---
title: Columns.Width property (Word)
keywords: vbawd10.chm155910147
f1_keywords:
- vbawd10.chm155910147
ms.prod: word
api_name:
- Word.Columns.Width
ms.assetid: 011c3c8f-1d80-a7d1-3a05-f634779f158e
ms.date: 06/08/2017
localization_priority: Normal
---


# Columns.Width property (Word)

Returns or sets the width of the specified columns, in points. Read/write  **Long**.


## Syntax

_expression_.**Width**

_expression_ A variable that represents a '[Columns](Word.columns.md)' collection.


## Example

This example creates a 5x5 table in a new document and then sets the width of all the columns in the table to 1.5 inches.


```vb
Dim objDoc As Document 
Dim objTable As Table 
 
Set objDoc = ActiveDocument 
Set objTable = objDoc.Tables.Add(Range:=Selection.Range, _ 
 NumRows:=5, NumColumns:=5) 
objTable.Columns.Width = InchesToPoints(1.5)
```


## See also


[Columns Collection Object](Word.columns.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
