---
title: CellRange Object (Publisher)
keywords: vbapb10.chm5242879
f1_keywords:
- vbapb10.chm5242879
ms.prod: publisher
api_name:
- Publisher.CellRange
ms.assetid: 86e164f3-2a04-013f-3da8-d45c013eae7b
ms.date: 06/08/2017
---


# CellRange Object (Publisher)

A collection of  **[Cell](Publisher.Cell.md)** objects in a table column or row. The **CellRange** collection represents all the cells in the specified column or row.
 


## Remarks

Although the collection object is named  **CellRange** and is shown in the Object Browser, this keyword is not used in programming the Microsoft Publisher object model. The keyword **Cells** is used instead.
 

 
You cannot programmatically add to or delete individual cells from a Publisher table. Use the  **[AddTable](Publisher.Shapes.AddTable.md)** method with the **[Shapes](Publisher.Shapes.md)** collection to add a new table to a publication. Use the **[Add](Publisher.Columns.Add.md)** method of the **[Columns](Publisher.Columns.md)** or **[Rows](Publisher.Rows.md)** collections to add a column or row to a table. Use the **[Delete](Publisher.Column.Delete.md)** method of the **Columns** or **Rows** collections to delete a column or row from a table.
 

 

## Example

Use the  **[Cells](Publisher.Column.Cells.md)** property to return the **CellRange** collection. This example merges the cells in first column of the table.
 

 

```
Sub MergeCellsInFirstColumn() 
 With ActiveDocument.Pages(1).Shapes(1).Table 
 .Cells(StartRow:=1, StartColumn:=1, _ 
 EndRow:=.Rows.Count, EndColumn:=1).Select 
 End With 
 Selection.TableCellRange.Merge 
End Sub
```

Use the  **[Count](Publisher.CellRange.Count.md)** property to return the number of cells in a row, column, table or selection. This example displays a message with the number of cells the specified table.
 

 



```
Sub NumberOfTableCells() 
 MsgBox ActiveDocument.Pages(1).Shapes(1).Table _ 
 .Cells.Count 
End Sub
```


## Methods



|**Name**|
|:-----|
|[Item](Publisher.CellRange.Item.md)|
|[Merge](Publisher.CellRange.Merge.md)|
|[Select](Publisher.CellRange.Select.md)|

## Properties



|**Name**|
|:-----|
|[Application](Publisher.CellRange.Application.md)|
|[Column](Publisher.CellRange.Column.md)|
|[Count](Publisher.CellRange.Count.md)|
|[Height](Publisher.CellRange.Height.md)|
|[Parent](Publisher.CellRange.Parent.md)|
|[Row](Publisher.CellRange.Row.md)|
|[Width](cellrange-width-property-publisher.md)|

