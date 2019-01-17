---
title: Column.Cells Property (Publisher)
keywords: vbapb10.chm4980738
f1_keywords:
- vbapb10.chm4980738
ms.prod: publisher
api_name:
- Publisher.Column.Cells
ms.assetid: 6c8b33f9-61f0-086c-1ceb-996221aa3a02
ms.date: 06/08/2017
localization_priority: Normal
---


# Column.Cells Property (Publisher)

Returns a  **[CellRange](Publisher.CellRange.md)** object that represents the cell or cells in a column of a table.


## Syntax

 _expression_. **Cells**

 _expression_ A variable that represents a  **Column** object.


## Example

This example merges the first and second cells in the first column of the specified table.


```vb
Sub MergeCell() 
 With ActiveDocument.Pages(1).Shapes(2).Table.Columns(1) 
 .Cells(1).Merge MergeTo:=.Cells(2) 
 End With 
End Sub
```

This example applies a thick border outline to the first cell in the second column of the specified table.




```vb
Sub OutlineBorderCell() 
 With ActiveDocument.Pages(1).Shapes(2).Table.Columns(2).Cells(1) 
 .BorderLeft.Weight = 5 
 .BorderRight.Weight = 5 
 .BorderTop.Weight = 5 
 .BorderBottom.Weight = 5 
 End With 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]