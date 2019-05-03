---
title: PivotCell.PivotCellType property (Excel)
keywords: vbaxl10.chm692073
f1_keywords:
- vbaxl10.chm692073
ms.prod: excel
api_name:
- Excel.PivotCell.PivotCellType
ms.assetid: f5462981-924c-4d6c-be99-5b7cea0222a4
ms.date: 05/04/2019
localization_priority: Normal
---


# PivotCell.PivotCellType property (Excel)

Returns one of the **[XlPivotCellType](Excel.XlPivotCellType.md)** constants that identifies the PivotTable entity that the cell corresponds to. Read-only.


## Syntax

_expression_.**PivotCellType**

_expression_ A variable that represents a **[PivotCell](Excel.PivotCell.md)** object.


## Example

This example determines if cell A5 in the PivotTable is a data item and notifies the user. The example assumes that a PivotTable exists on the active worksheet, and cell A5 is contained in the PivotTable. If cell A5 is not in the PivotTable, the example handles the run-time error.

```vb
Sub CheckPivotCellType() 
 
 On Error GoTo Not_In_PivotTable 
 
 ' Determine if cell A5 is a data item in the PivotTable. 
 If Application.Range("A5").PivotCell.PivotCellType = xlPivotCellValue Then 
 MsgBox "The cell at A5 is a data item." 
 Else 
 MsgBox "The cell at A5 is not a data item." 
 End If 
 Exit Sub 
 
Not_In_PivotTable: 
 MsgBox "The chosen cell is not in a PivotTable." 
 
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]