---
title: Rows.DistributeHeight method (Word)
keywords: vbawd10.chm155975886
f1_keywords:
- vbawd10.chm155975886
ms.prod: word
api_name:
- Word.Rows.DistributeHeight
ms.assetid: f5fe9eea-debc-c1e4-b9a0-81c5f9a0c04a
ms.date: 06/08/2017
localization_priority: Normal
---


# Rows.DistributeHeight method (Word)

Adjusts the height of the specified rows or cells so that they're equal.


## Syntax

_expression_. `DistributeHeight`

_expression_ Required. A variable that represents a **[Rows](Word.Rows.md)** object.


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


[Rows Collection Object](Word.rows.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]