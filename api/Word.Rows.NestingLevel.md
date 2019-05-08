---
title: Rows.NestingLevel property (Word)
keywords: vbawd10.chm155975783
f1_keywords:
- vbawd10.chm155975783
ms.prod: word
api_name:
- Word.Rows.NestingLevel
ms.assetid: 54a34d92-08bc-fb66-3a29-5e491d370307
ms.date: 06/08/2017
localization_priority: Normal
---


# Rows.NestingLevel property (Word)

Returns the nesting level of the specified table rows. Read-only  **Long**.


## Syntax

_expression_. `NestingLevel`

_expression_ Required. A variable that represents a **[Rows](Word.Rows.md)** object.


## Remarks

The outermost table has a nesting level of 1. The nesting level of each successively nested table is one higher than the previous table.


## Example

This example creates a new document, creates a nested table with three levels, and then fills the first cell of each table with its nesting level.


```vb
Documents.Add 
ActiveDocument.Tables.Add Selection.Range, _ 
 3, 3, wdWord9TableBehavior, wdAutoFitContent 
With ActiveDocument.Tables(1).Range 
 .Copy 
 .Cells(1).Range.Text = .Cells(1).NestingLevel 
 .Cells(5).Range.PasteAsNestedTable 
 With .Cells(5).Tables(1).Range 
 .Cells(1).Range.Text = .Cells(1).NestingLevel 
 .Cells(5).Range.PasteAsNestedTable 
 With .Cells(5).Tables(1).Range 
 .Cells(1).Range.Text = _ 
 .Cells(1).NestingLevel 
 End With 
 End With 
End With
```


## See also


[Rows Collection Object](Word.rows.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]