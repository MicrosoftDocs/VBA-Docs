---
title: CellRange object (Publisher)
keywords: vbapb10.chm5242879
f1_keywords:
- vbapb10.chm5242879
ms.prod: publisher
api_name:
- Publisher.CellRange
ms.assetid: 86e164f3-2a04-013f-3da8-d45c013eae7b
ms.date: 05/31/2019
localization_priority: Normal
---


# CellRange object (Publisher)

A collection of **[Cell](Publisher.Cell.md)** objects in a table column or row. The **CellRange** collection represents all the cells in the specified column or row.
 

## Remarks

Although the collection object is named **CellRange** and is shown in the Object Browser, this keyword is not used in programming the Microsoft Publisher object model. The keyword **Cells** is used instead.

You cannot programmatically add to or delete individual cells from a Publisher table: 

- Use the **[AddTable](Publisher.Shapes.AddTable.md)** method of the **Shapes** collection to add a new table to a publication. 
- Use the **[Add](Publisher.Columns.Add.md)** method of the **Columns** or **[Rows](Publisher.Rows.md)** collections to add a column or row to a table. 
- Use the **[Delete](Publisher.Column.Delete.md)** method of the **Column** or **[Row](Publisher.Row.md)** objects to delete a column or row from a table.

Use the **[Cells](Publisher.Column.Cells.md)** property of the **Column** object to return the **CellRange** collection.

Use the **Count** property to return the number of cells in a row, column, table, or selection. 

## Example

This example merges the cells in the first column of the table.

```vb
Sub MergeCellsInFirstColumn() 
 With ActiveDocument.Pages(1).Shapes(1).Table 
 .Cells(StartRow:=1, StartColumn:=1, _ 
 EndRow:=.Rows.Count, EndColumn:=1).Select 
 End With 
 Selection.TableCellRange.Merge 
End Sub
```

<br/>

This example displays a message with the number of cells in the specified table.

```vb
Sub NumberOfTableCells() 
 MsgBox ActiveDocument.Pages(1).Shapes(1).Table _ 
 .Cells.Count 
End Sub
```


## Methods

- [Item](Publisher.CellRange.Item.md)
- [Merge](Publisher.CellRange.Merge.md)
- [Select](Publisher.CellRange.Select.md)

## Properties

- [Application](Publisher.CellRange.Application.md)
- [Column](Publisher.CellRange.Column.md)
- [Count](Publisher.CellRange.Count.md)
- [Height](Publisher.CellRange.Height.md)
- [Parent](Publisher.CellRange.Parent.md)
- [Row](Publisher.CellRange.Row.md)
- [Width](Publisher.CellRange.Width.md)

## See also

- [Publisher Object Model Reference](overview/publisher/object-model.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]