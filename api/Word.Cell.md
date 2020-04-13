---
title: Cell object (Word)
keywords: vbawd10.chm2382
f1_keywords:
- vbawd10.chm2382
ms.prod: word
api_name:
- Word.Cell
ms.assetid: cbe6ae71-b2da-63a9-1446-0a2f81ab8b14
ms.date: 06/08/2017
localization_priority: Normal
---


# Cell object (Word)

Represents a single table cell. The **Cell** object is a member of the **[Cells](Word.cells.md)** collection. The **Cells** collection represents all the cells in the specified object.


## Remarks

Use  **[Cell](Word.Table.Cell.md)** (row, column), where row is the row number and column is the column number, or **Cells** (_index_), where _index_ is the index number, to return a **Cell** object. The following example applies shading to the second cell in the first row.


```vb
Set myCell = ActiveDocument.Tables(1).Cell(Row:=1, Column:=2) 
myCell.Shading.Texture = wdTexture20Percent
```

The following example applies shading to the first cell in the first row.




```vb
ActiveDocument.Tables(1).Rows(1).Cells(1).Shading _ 
 .Texture = wdTexture20Percent
```

Use the **[Add](Word.Cells.Add.md)** method to add a **Cell** object to the **[Cells](Word.cells.md)** collection. You can also use the **[InsertCells](Word.Selection.InsertCells.md)** method of the **[Selection](Word.Selection.md)** object to insert new cells. The following example adds a cell before the first cell in `myTable`.




```vb
Set myTable = ActiveDocument.Tables(1) 
myTable.Range.Cells.Add BeforeCell:=myTable.Cell(1, 1)
```

The following example sets a range ( _myRange_ ) that references the first two cells in the first table. After the range is set, the cells are combined by the **[Merge](Word.Cells.Merge.md)** method.




```vb
Set myTable = ActiveDocument.Tables(1) 
Set myRange = ActiveDocument.Range(myTable.Cell(1, 1) _ 
 .Range.Start, myTable.Cell(1, 2).Range.End) 
myRange.Cells.Merge
```

Remarks

Use the **[Add](Word.AddIns.Add.md)** method with the **[Rows](Word.rows.md)** or **[Columns](Word.columns.md)** collection to add a row or column of cells.

Use the **[Information](Word.Selection.Information.md)** property with a **Selection** object to return the current row and column number. The following example changes the width of the first cell in the selection and then displays the cell's row number and column number.




```vb
If Selection.Information(wdWithInTable) = True Then 
 With Selection 
 .Cells(1).Width = 22 
 MsgBox "Cell " & .Information(wdStartOfRangeRowNumber) _ 
 & "," & .Information(wdStartOfRangeColumnNumber) 
 End With 
End If
```


## See also


[Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
