---
title: PivotCell object (Excel)
keywords: vbaxl10.chm691072
f1_keywords:
- vbaxl10.chm691072
ms.prod: excel
api_name:
- Excel.PivotCell
ms.assetid: 76b8a2dc-90ee-7475-d327-d27cb1e92703
ms.date: 06/08/2017
localization_priority: Normal
---


# PivotCell object (Excel)

Represents a cell in a PivotTable report.


## Remarks

Use the  **[PivotCell](Excel.Range.PivotCell.md)** property of the **[Range](Excel.Range(object).md)** collection to return a **PivotCell** object.

Once a  **PivotCell** object is returned, you can use the **[ColumnItems](Excel.PivotCell.ColumnItems.md)** or **[RowItems](Excel.PivotCell.RowItems.md)** property to determine the **[PivotItems](Excel.PivotItems.md)** collection that corresponds to the items on the column or row axis that represents the selected number. The following example uses the **ColumnItems** property of the **PivotCell** object to return a **[PivotItemList](Excel.PivotItemList.md)** collection.


## Example

Once a  **PivotCell** object is returned, you can use the **[PivotCellType](Excel.PivotCell.PivotCellType.md)** property to determine what type of cell a particular range is. The following example determines if cell A5 in the PivotTable is a data item and notifies the user. This example assumes that a PivotTable exists on the active worksheet and that cell A5 is contained in the PivotTable. If cell A5 is not in the PivotTable, the example handles the run-time error.


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

This example determines the column field that the data item of cell B5 is in. It then determines if the column field title matches "Inventory" and notifies the user. The example assumes that a PivotTable exists on the active worksheet and that column B of the worksheet contains a column field of the PivotTable.




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


## See also


[Excel Object Model Reference](overview/Excel/object-model.md)


