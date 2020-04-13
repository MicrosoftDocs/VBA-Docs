---
title: Cells object (Word)
ms.prod: word
ms.assetid: ceaa5b45-518d-d6ea-1ce8-5a34f6e37046
ms.date: 06/08/2017
localization_priority: Normal
---


# Cells object (Word)

A collection of  **[Cell](Word.Cell.md)** objects in a table column, table row, selection, or range.


## Remarks

Use the **Cells** property to return the **Cells** collection. The following example formats the cells in the first row in table one in the active document to be 30 points wide.


```vb
ActiveDocument.Tables(1).Rows(1).Cells.Width = 30
```

The following example returns the number of cells in the current row.




```vb
num = Selection.Rows(1).Cells.Count
```

Use the **[Add](Word.Cells.Add.md)** method to add a **[Cell](Word.Cell.md)** object to the **Cells** collection. You can also use the **[InsertCells](Word.Selection.InsertCells.md)** method of the **[Selection](Word.Selection.md)** object to insert new cells. The following example adds a cell before the first cell in myTable.




```vb
Set myTable = ActiveDocument.Tables(1) 
myTable.Range.Cells.Add BeforeCell:=myTable.Cell(1, 1)
```

Use  **Cell** (row, column), where row is the row number and column is the column number, or **Cells** (_index_), where _index_ is the index number, to return a **Cell** object. The following example applies shading to the second cell in the first row in table one.




```vb
Set myCell = ActiveDocument.Tables(1).Cell(Row:=1, Column:=2) 
myCell.Shading.Texture = wdTexture20Percent
```

The following example applies shading to the first cell in the first row.




```vb
ActiveDocument.Tables(1).Rows(1).Cells(1).Shading _ 
 .Texture = wdTexture20Percent
```

Remarks

Use the **Add** method with the **[Rows](Word.rows.md)** or **[Columns](Word.columns.md)** collection to add a row or column of cells. The following example adds a column to the first table in the active document and then inserts numbers into the first column.




```vb
Set myTable = ActiveDocument.Tables(1) 
Set aColumn = myTable.Columns.Add(BeforeColumn:=myTable.Columns(1)) 
For Each aCell In aColumn.Cells 
 aCell.Range.Delete 
 aCell.Range.InsertAfter num + 1 
 num = num + 1 
Next aCell
```


## Methods



|Name|
|:-----|
|[Add](Word.Cells.Add.md)|
|[AutoFit](Word.Cells.AutoFit.md)|
|[Delete](Word.Cells.Delete.md)|
|[DistributeHeight](Word.Cells.DistributeHeight.md)|
|[DistributeWidth](Word.Cells.DistributeWidth.md)|
|[Item](Word.Cells.Item.md)|
|[Merge](Word.Cells.Merge.md)|
|[SetHeight](Word.Cells.SetHeight.md)|
|[SetWidth](Word.Cells.SetWidth.md)|
|[Split](Word.Cells.Split.md)|

## Properties



|Name|
|:-----|
|[Application](Word.Cells.Application.md)|
|[Borders](Word.Cells.Borders.md)|
|[Count](Word.Cells.Count.md)|
|[Creator](Word.Cells.Creator.md)|
|[Height](Word.Cells.Height.md)|
|[HeightRule](Word.Cells.HeightRule.md)|
|[NestingLevel](Word.Cells.NestingLevel.md)|
|[Parent](Word.Cells.Parent.md)|
|[PreferredWidth](Word.Cells.PreferredWidth.md)|
|[PreferredWidthType](Word.Cells.PreferredWidthType.md)|
|[Shading](Word.Cells.Shading.md)|
|[VerticalAlignment](Word.Cells.VerticalAlignment.md)|
|[Width](Word.Cells.Width.md)|

## See also


[Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
