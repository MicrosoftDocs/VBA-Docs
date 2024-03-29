---
title: CellFormat.Borders property (Excel)
keywords: vbaxl10.chm676073
f1_keywords:
- vbaxl10.chm676073
api_name:
- Excel.CellFormat.Borders
ms.assetid: 8a2ad449-a1b4-14ff-6a67-f475dba82c45
ms.date: 04/16/2019
ms.localizationpriority: medium
---


# CellFormat.Borders property (Excel)

Returns or sets a **[Borders](Excel.Borders.md)** collection that represents the search criteria based on the cell's border format.


## Syntax

_expression_.**Borders**

_expression_ A variable that represents a **[CellFormat](Excel.CellFormat.md)** object.


## Example

This example sets the search criteria to identify the borders of cells that have a continuous and thick style bottom-edge, creates a cell with this condition, finds this cell, and notifies the user. 

> [!NOTE] 
> The default color of the border is used in this example; therefore, the color index is not changed.

```vb
Sub SearchCellFormat() 
 
 ' Set the search criteria for the border of the cell format. 
 With Application.FindFormat.Borders(xlEdgeBottom) 
 .LineStyle = xlContinuous 
 .Weight = xlThick 
 End With 
 
 ' Create a continuous thick bottom-edge border for cell A5. 
 Range("A5").Select 
 With Selection.Borders(xlEdgeBottom) 
 .LineStyle = xlContinuous 
 .Weight = xlThick 
 End With 
 Range("A1").Select 
 MsgBox "Cell A5 has a continuous thick bottom-edge border" 
 
 ' Find the cells based on the search criteria. 
 Cells.Find(What:="", After:=ActiveCell, LookIn:=xlFormulas, LookAt:= _ 
 xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _ 
 , SearchFormat:=True).Activate 
 MsgBox "Microsoft Excel has found this cell matching the search criteria." 
 
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]