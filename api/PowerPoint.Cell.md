---
title: Cell object (PowerPoint)
keywords: vbapp10.chm628000
f1_keywords:
- vbapp10.chm628000
ms.prod: powerpoint
api_name:
- PowerPoint.Cell
ms.assetid: e89e5d69-33b1-d7b1-0a6c-4dfd8b676977
ms.date: 06/08/2017
localization_priority: Normal
---


# Cell object (PowerPoint)

Represents a table cell. The  **Cell** object is a member of the **[CellRange](PowerPoint.CellRange.md)** collection. The **CellRange** collection represents all the cells in the specified column or row. To use the **CellRange** collection, use the **Cells** keyword.


## Remarks

You cannot programmatically add cells to or delete cells from a PowerPoint table. Use the  **Add** method of the **Columns** or **Rows** collections to add a column or row to a table. Use the **Delete** method of the **Columns** or **Rows** collections to delete a column or row from a table.


## Example

Use  **Cell** (row, column), where row is the row number and column is the column number, or **Cells** (_index_), where _index_ is the number of the cell in the specified row or column, to return a single **Cell** object. Cells are numbered from left to right in rows and from top to bottom in columns. With right-to-left language settings, this scheme is reversed. The following example merges the first two cells in row one of the table in shape five on slide two.


```vb
With ActivePresentation.Slides(2).Shapes(5).Table

    .Cell(1, 1).Merge MergeTo:=.Cell(1, 2)

End With
```

This example sets the bottom border for cell one in the first column of the table to a dashed line style.




```vb
With ActivePresentation.Slides(2).Shapes(5).Table.Columns(1) _

        .Cells(1)

    .Borders(ppBorderBottom).DashStyle = msoLineDash

End With
```

Use the [Shape](PowerPoint.Cell.Shape.md)property to access the  **Shape** object and to manipulate the contents of each cell. This example deletes the text in the first cell (row 1, column 1), inserts new text, and then sets the width of the entire column to 110 points.




```vb
With ActivePresentation.Slides(2).Shapes(5).Table.Cell(1, 1)

    .Shape.TextFrame.TextRange.Delete

    .Shape.TextFrame.TextRange.Text = "Rooster"

    .Parent.Columns(1).Width = 110

End With
```


## Methods



|Name|
|:-----|
|[Merge](PowerPoint.Cell.Merge.md)|
|[Select](PowerPoint.Cell.Select.md)|
|[Split](PowerPoint.Cell.Split.md)|

## Properties



|Name|
|:-----|
|[Application](PowerPoint.Cell.Application.md)|
|[Borders](PowerPoint.Cell.Borders.md)|
|[Parent](PowerPoint.Cell.Parent.md)|
|[Selected](PowerPoint.Cell.Selected.md)|
|[Shape](PowerPoint.Cell.Shape.md)|

## See also


[PowerPoint Object Model Reference](overview/PowerPoint/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]