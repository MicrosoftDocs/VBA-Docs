---
title: Protection object (Excel)
keywords: vbaxl10.chm719072
f1_keywords:
- vbaxl10.chm719072
ms.prod: excel
api_name:
- Excel.Protection
ms.assetid: dc13a9dd-bd19-daa2-5093-7182917d5bde
ms.date: 03/30/2019
localization_priority: Normal
---


# Protection object (Excel)

Represents the various types of protection options available for a worksheet.


## Remarks

Use the **[Protection](Excel.Worksheet.Protection.md)** property of the **Worksheet** object to return a **Protection** object.

After a **Protection** object is returned, you can use the **[Protection](#properties)** properties to set or return protection options.

## Example

The following example demonstrates how to use the **AllowInsertingColumns** property of the **Protection** object, placing three numbers in the top row and protecting the worksheet. This example then checks to see if the protection setting for allowing the insertion of columns is **False** and sets it to **True**, if necessary. Finally, it notifies the user to insert a column.

```vb
Sub SetProtection() 
 
 Range("A1").Formula = "1" 
 Range("B1").Formula = "3" 
 Range("C1").Formula = "4" 
 ActiveSheet.Protect 
 
 ' Check the protection setting of the worksheet and act accordingly. 
 If ActiveSheet.Protection.AllowInsertingColumns = False Then 
 ActiveSheet.Protect AllowInsertingColumns:=True 
 MsgBox "Insert a column between 1 and 3" 
 Else 
 MsgBox "Insert a column between 1 and 3" 
 End If 
 
End Sub
```


## Properties

- [AllowDeletingColumns](Excel.Protection.AllowDeletingColumns.md)
- [AllowDeletingRows](Excel.Protection.AllowDeletingRows.md)
- [AllowEditRanges](Excel.Protection.AllowEditRanges.md)
- [AllowFiltering](Excel.Protection.AllowFiltering.md)
- [AllowFormattingCells](Excel.Protection.AllowFormattingCells.md)
- [AllowFormattingColumns](Excel.Protection.AllowFormattingColumns.md)
- [AllowFormattingRows](Excel.Protection.AllowFormattingRows.md)
- [AllowInsertingColumns](Excel.Protection.AllowInsertingColumns.md)
- [AllowInsertingHyperlinks](Excel.Protection.AllowInsertingHyperlinks.md)
- [AllowInsertingRows](Excel.Protection.AllowInsertingRows.md)
- [AllowSorting](Excel.Protection.AllowSorting.md)
- [AllowUsingPivotTables](Excel.Protection.AllowUsingPivotTables.md)


## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
