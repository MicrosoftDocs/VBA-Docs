---
title: Cell Object (Publisher)
keywords: vbapb10.chm5177343
f1_keywords:
- vbapb10.chm5177343
ms.prod: publisher
api_name:
- Publisher.Cell
ms.assetid: 5baafaa6-368e-9eae-30b9-90d2d89d5a5b
ms.date: 06/08/2017
---


# Cell Object (Publisher)

Represents a single table cell. The  **Cell** object is a member of the **[CellRange](./Publisher.CellRange.md)** collection. The **CellRange** collection represents all the cells in the specified object.


## Example

Use  **Cells** (index), where index is the cell number, to return a **Cell** object. This example merges the first two cells of the first column of the specified table.


```vb
Sub MergeCell() 
 With ActiveDocument.Pages(1).Shapes(2).Table.Columns(1) 
 .Cells(1).Merge MergeTo:=.Cells(2) 
 End With 
End Sub
```

This example applies a thick border around the first cell in the second column of the specified table.




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


## Methods



|**Name**|
|:-----|
|[Merge](./Publisher.Cell.Merge.md)|
|[Select](./Publisher.Cell.Select.md)|
|[Split](./Publisher.Cell.Split.md)|

## Properties



|**Name**|
|:-----|
|[Application](./Publisher.Cell.Application.md)|
|[BorderBottom](./Publisher.Cell.BorderBottom.md)|
|[BorderDiagonal](./Publisher.Cell.BorderDiagonal.md)|
|[BorderLeft](./Publisher.Cell.BorderLeft.md)|
|[BorderRight](./Publisher.Cell.BorderRight.md)|
|[BorderTop](./Publisher.Cell.BorderTop.md)|
|[CellTextOrientation](./Publisher.Cell.CellTextOrientation.md)|
|[Column](./Publisher.Cell.Column.md)|
|[Diagonal](./Publisher.Cell.Diagonal.md)|
|[Fill](./Publisher.Cell.Fill.md)|
|[HasText](./Publisher.Cell.HasText.md)|
|[Height](./Publisher.Cell.Height.md)|
|[MarginBottom](./Publisher.Cell.MarginBottom.md)|
|[MarginLeft](./Publisher.Cell.MarginLeft.md)|
|[MarginRight](./Publisher.Cell.MarginRight.md)|
|[MarginTop](./Publisher.Cell.MarginTop.md)|
|[Parent](./Publisher.Cell.Parent.md)|
|[Row](./Publisher.Cell.Row.md)|
|[Selected](./Publisher.Cell.Selected.md)|
|[TextRange](./Publisher.Cell.TextRange.md)|
|[VerticalTextAlignment](./Publisher.Cell.VerticalTextAlignment.md)|
|[Width](./Publisher.Cell.Width.md)|

