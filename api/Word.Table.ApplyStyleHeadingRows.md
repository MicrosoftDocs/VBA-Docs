---
title: Table.ApplyStyleHeadingRows property (Word)
keywords: vbawd10.chm156303562
f1_keywords:
- vbawd10.chm156303562
ms.prod: word
api_name:
- Word.Table.ApplyStyleHeadingRows
ms.assetid: 1c7fb6d5-9010-fded-d882-388d1e631da2
ms.date: 06/08/2017
localization_priority: Normal
---


# Table.ApplyStyleHeadingRows property (Word)

 **True** for Microsoft Word to apply heading-row formatting to the first row of the selected table. Read/write **Boolean**.


## Syntax

_expression_. `ApplyStyleHeadingRows`

 _expression_ An expression that returns a '[Table](Word.Table.md)' object.


## Remarks

The specified table style must contain heading-row formatting to apply this formatting to a table.


## Example

This example formats the second table in the active document with the table style "Table Style 1" and removes formatting for the first and last rows and the first and last columns. This example assumes that a table style named "Table Style 1" exists and that it contains heading-row formatting.


```vb
Sub TableStyles() 
 With ActiveDocument.Tables(2) 
 .Style = "Table Style 1" 
 .ApplyStyleFirstColumn = False 
 .ApplyStyleHeadingRows = False 
 .ApplyStyleLastColumn = False 
 .ApplyStyleLastRow = False 
 End With 
End Sub
```


## See also


[Table Object](Word.Table.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]