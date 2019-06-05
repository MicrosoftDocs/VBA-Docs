---
title: Cell.BorderDiagonal property (Publisher)
keywords: vbapb10.chm5111810
f1_keywords:
- vbapb10.chm5111810
ms.prod: publisher
api_name:
- Publisher.Cell.BorderDiagonal
ms.assetid: 2c857a1b-2a0f-5796-9397-ad113dd984cb
ms.date: 06/06/2019
localization_priority: Normal
---


# Cell.BorderDiagonal property (Publisher)

Returns a **[CellBorder](Publisher.CellBorder.md)** object that represents the diagonal border for a specified table cell.


## Syntax

_expression_.**BorderDiagonal**

_expression_ A variable that represents a **[Cell](Publisher.Cell.md)** object.


## Return value

CellBorder


## Example

This example diagonally splits every other cell in the specified table and adds a diagonal border. This example assumes that the first shape on page two is a table and not another type of shape.

```vb
Sub FillCellsByRow() 
 Dim shpTable As Shape 
 Dim rowTable As Row 
 Dim celTable As Cell 
 Dim intCell As Integer 
 
 intCell = 1 
 
 Set shpTable = ActiveDocument.Pages(2).Shapes(1) 
 For Each rowTable In shpTable.Table.Rows 
 For Each celTable In rowTable.Cells 
 If intCell Mod 2 = 0 Then 
 With celTable 
 .Diagonal = pbTableCellDiagonalDown 
 With .BorderDiagonal 
 .Weight = 1 
 .Color.RGB = RGB(Red:=0, Green:=0, Blue:=0) 
 End With 
 End With 
 celTable.Fill.ForeColor.RGB = RGB _ 
 (Red:=180, Green:=180, Blue:=180) 
 Else 
 celTable.Fill.ForeColor.RGB = RGB _ 
 (Red:=255, Green:=255, Blue:=255) 
 End If 
 intCell = intCell + 1 
 Next celTable 
 Next rowTable 
 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]