---
title: Cell.Merge method (Publisher)
keywords: vbapb10.chm5111842
f1_keywords:
- vbapb10.chm5111842
ms.prod: publisher
api_name:
- Publisher.Cell.Merge
ms.assetid: 09ae6910-ba47-4b0e-5792-ac9eb1298d57
ms.date: 06/06/2019
localization_priority: Normal
---


# Cell.Merge method (Publisher)

Merges the specified table cell with another cell. The result is a single table cell.


## Syntax

_expression_.**Merge** (_MergeTo_)

_expression_ A variable that represents a **[Cell](Publisher.Cell.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_MergeTo_|Required| **Cell**|The cell to be merged with; must be adjacent to the specified cell or an error occurs.|


## Example

This example merges the first two cells of the first column of the specified table.

```vb
Sub MergeCell() 
 With ActiveDocument.Pages(1).Shapes(2).Table 
 .Rows(1).Cells(1).Merge MergeTo:=.Rows(2).Cells(1) 
 End With 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]