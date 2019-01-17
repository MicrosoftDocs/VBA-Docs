---
title: Cells.DistributeHeight method (Word)
keywords: vbawd10.chm155844814
f1_keywords:
- vbawd10.chm155844814
ms.prod: word
api_name:
- Word.Cells.DistributeHeight
ms.assetid: 0ae41e05-5ec1-4fcc-8ee1-c40c0a28714a
ms.date: 06/08/2017
localization_priority: Normal
---


# Cells.DistributeHeight method (Word)

Adjusts the height of the specified cells so that they are equal.


## Syntax

 _expression_. `DistributeHeight`

 _expression_ Required. A variable that represents a '[Cells](Word.cells.md)' collection.


## Example

This example adjusts the height of the rows in the first table in the active document so that they're equal.


```vb
ActiveDocument.Tables(1).Rows.DistributeHeight
```

This example adjusts the height of the first three rows in the first table so that they're equal.




```vb
Dim rngTemp As Range 
 
If ActiveDocument.Tables.Count >= 1 Then 
 Set rngTemp = ActiveDocument.Range(Start:=ActiveDocument _ 
 .Tables(1).Rows(1).Range.Start, _ 
 End:=ActiveDocument.Tables(1).Rows(3).Range.End) 
 rngTemp.Rows.DistributeHeight 
End If
```


## See also


[Cells Collection Object](Word.cells.md)

