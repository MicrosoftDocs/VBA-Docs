---
title: Cell.Diagonal property (Publisher)
keywords: vbapb10.chm5111816
f1_keywords:
- vbapb10.chm5111816
ms.prod: publisher
api_name:
- Publisher.Cell.Diagonal
ms.assetid: 4ec93690-38ef-7434-55a5-419f14c9ea73
ms.date: 06/06/2019
localization_priority: Normal
---


# Cell.Diagonal property (Publisher)

Sets or returns a **[PbCellDiagonalType](Publisher.PbCellDiagonalType.md)** constant that represents a cell that is diagonally split. Read/write.


## Syntax

_expression_.**Diagonal**

_expression_ A variable that represents a **[Cell](Publisher.Cell.md)** object.


## Return value

PbCellDiagonalType


## Remarks

The **Diagonal** property value can be one of the **PbCellDiagonalType** constants declared in the Microsoft Publisher type library.


## Example

This example adds a page to the active publication, creates a table on that new page, and diagonally splits all cells in even-numbered columns.

```vb
Sub CreateNewTable() 
 
 Dim pgeNew As Page 
 Dim shpTable As Shape 
 Dim tblNew As Table 
 Dim celTable As Cell 
 Dim rowTable As Row 
 
 'Creates a new document with a five-row by five-column table 
 Set pgeNew = ActiveDocument.Pages.Add(Count:=1, After:=1) 
 Set shpTable = pgeNew.Shapes.AddTable(NumRows:=5, NumColumns:=5, _ 
 Left:=72, Top:=72, Width:=468, Height:=100) 
 Set tblNew = shpTable.Table 
 
 'Inserts a diagonal split into all cells in even-numbered columns 
 For Each rowTable In tblNew.Rows 
 For Each celTable In rowTable.Cells 
 If celTable.Column Mod 2 = 0 Then 
 celTable.Diagonal = pbTableCellDiagonalUp 
 End If 
 Next celTable 
 Next rowTable 
 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]