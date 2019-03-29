---
title: PivotCell object (Excel)
keywords: vbaxl10.chm691072
f1_keywords:
- vbaxl10.chm691072
ms.prod: excel
api_name:
- Excel.PivotCell
ms.assetid: 76b8a2dc-90ee-7475-d327-d27cb1e92703
ms.date: 03/30/2019
localization_priority: Normal
---


# PivotCell object (Excel)

Represents a cell in a PivotTable report.


## Remarks

Use the **[PivotCell](Excel.Range.PivotCell.md)** property of the **Range** collection to return a **PivotCell** object.

After a **PivotCell** object is returned, you can use the **ColumnItems** or **RowItems** property to determine the **[PivotItems](Excel.PivotItems.md)** collection that corresponds to the items on the column or row axis that represents the selected number. 

## Example

After a **PivotCell** object is returned, you can use the **PivotCellType** property to determine what type of cell a particular range is. 

The following example determines if cell A5 in the PivotTable is a data item and notifies the user. This example assumes that a PivotTable exists on the active worksheet, and that cell A5 is contained in the PivotTable. If cell A5 is not in the PivotTable, the example handles the run-time error.

```vb
Sub CheckPivotCellType() 
 
 On Error GoTo Not_In_PivotTable 
 
 ' Determine if cell A5 is a data item in the PivotTable. 
 If Application.Range("A5").PivotCell.PivotCellType = xlPivotCellValue Then 
 MsgBox "The PivotCell at A5 is a data item." 
 Else 
 MsgBox "The PivotCell at A5 is not a data item." 
 End If 
 Exit Sub 
 
Not_In_PivotTable: 
 MsgBox "The chosen cell is not in a PivotTable." 
 
End Sub
```

<br/>

This example determines the column field that the data item of cell B5 is in. It then determines if the column field title matches "Inventory" and notifies the user. The example assumes that a PivotTable exists on the active worksheet, and that column B of the worksheet contains a column field of the PivotTable.

```vb
Sub CheckColumnItems() 
 
 ' Determine if there is a match between the item and column field. 
 If Application.Range("B5").PivotCell.ColumnItems.Item(1) = "Inventory" Then 
 MsgBox "Item in B5 is a member of the 'Inventory' column field." 
 Else 
 MsgBox "Item in B5 is not a member of the 'Inventory' column field." 
 End If 
 
End Sub
```

## Methods

- [AllocateChange](Excel.PivotCell.AllocateChange.md)
- [DiscardChange](Excel.PivotCell.DiscardChange.md)

## Properties

- [Application](Excel.PivotCell.Application.md)
- [CellChanged](Excel.PivotCell.CellChanged.md)
- [ColumnItems](Excel.PivotCell.ColumnItems.md)
- [Creator](Excel.PivotCell.Creator.md)
- [CustomSubtotalFunction](Excel.PivotCell.CustomSubtotalFunction.md)
- [DataField](Excel.PivotCell.DataField.md)
- [DataSourceValue](Excel.PivotCell.DataSourceValue.md)
- [MDX](Excel.PivotCell.MDX.md)
- [Parent](Excel.PivotCell.Parent.md)
- [PivotCellType](Excel.PivotCell.PivotCellType.md)
- [PivotColumnLine](Excel.PivotCell.PivotColumnLine.md)
- [PivotField](Excel.PivotCell.PivotField.md)
- [PivotItem](Excel.PivotCell.PivotItem.md)
- [PivotRowLine](Excel.PivotCell.PivotRowLine.md)
- [PivotTable](Excel.PivotCell.PivotTable.md)
- [Range](Excel.PivotCell.Range.md)
- [RowItems](Excel.PivotCell.RowItems.md)
- [ServerActions](Excel.pivotcell.serveractions.md)

## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]