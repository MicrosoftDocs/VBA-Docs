---
title: DataTable.HasBorderOutline property (Word)
keywords: vbawd10.chm46399494
f1_keywords:
- vbawd10.chm46399494
ms.prod: word
api_name:
- Word.DataTable.HasBorderOutline
ms.assetid: c7766f52-ee4f-f51b-a716-b1b76dcb434f
ms.date: 06/08/2017
localization_priority: Normal
---


# DataTable.HasBorderOutline property (Word)

 **True** if the chart data table has outline borders. Read/write **Boolean**.


## Syntax

_expression_.**HasBorderOutline**

_expression_ A variable that represents a '[DataTable](Word.DataTable.md)' object.


## Example

The following example causes the data table for the first chart in the active document to be displayed with an outline border and no cell borders.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 With .Chart 
 .HasDataTable = True 
 With .DataTable 
 .HasBorderHorizontal = False 
 .HasBorderVertical = False 
 .HasBorderOutline = True 
 End With 
 End With 
 End If 
End With
```


## See also


[DataTable Object](Word.DataTable.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]