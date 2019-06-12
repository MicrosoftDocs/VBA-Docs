---
title: Row.Cells property (Publisher)
keywords: vbapb10.chm4849666
f1_keywords:
- vbapb10.chm4849666
ms.prod: publisher
api_name:
- Publisher.Row.Cells
ms.assetid: 2a866890-d564-b9bc-c553-06669f376788
ms.date: 06/13/2019
localization_priority: Normal
---


# Row.Cells property (Publisher)

Returns a **[CellRange](Publisher.CellRange.md)** object that represents the cell or cells in a row of a table.


## Syntax

_expression_.**Cells**

_expression_ A variable that represents a **[Row](Publisher.Row.md)** object.


## Example

This example merges the first and second cells in the first column of the specified table.

```vb
Sub MergeCell() 
 With ActiveDocument.Pages(1).Shapes(2).Table.Columns(1) 
 .Cells(1).Merge MergeTo:=.Cells(2) 
 End With 
End Sub
```

<br/>

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