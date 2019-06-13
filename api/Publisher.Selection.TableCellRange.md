---
title: Selection.TableCellRange property (Publisher)
keywords: vbapb10.chm851975
f1_keywords:
- vbapb10.chm851975
ms.prod: publisher
api_name:
- Publisher.Selection.TableCellRange
ms.assetid: d683e830-6bcd-4b53-844b-605fab184a4c
ms.date: 06/13/2019
localization_priority: Normal
---


# Selection.TableCellRange property (Publisher)

Returns a **[CellRange](publisher.cellrange.md)** object that represents the cells in a table selection.


## Syntax

_expression_.**TableCellRange**

_expression_ A variable that represents a **[Selection](Publisher.Selection.md)** object.


## Return value

CellRange


## Example

This example fills the table cells in a selection.

```vb
Sub FillTableCellRange() 
 Dim intCount As Integer 
 With Selection 
 If .Type = pbSelectionTableCells Then 
 With .TableCellRange 
 For intCount = 1 To .Count 
 .Item(intCount).Fill.ForeColor.RGB = RGB _ 
 (Red:=0, Green:=255, Blue:=255) 
 Next intCount 
 End With 
 End If 
 End With 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]