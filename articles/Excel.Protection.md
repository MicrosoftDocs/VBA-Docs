---
title: Protection Object (Excel)
keywords: vbaxl10.chm719072
f1_keywords:
- vbaxl10.chm719072
ms.prod: excel
api_name:
- Excel.Protection
ms.assetid: dc13a9dd-bd19-daa2-5093-7182917d5bde
ms.date: 06/08/2017
---


# Protection Object (Excel)

Represents the various types of protection options available for a worksheet.


## Remarks

Use the  **[Protection](Excel.Worksheet.Protection.md)** property of the **[Worksheet](Excel.Worksheet.md)** object to return a **Protection** object.

Once a  **Protection** object is returned, you can use its following properties, to set or return protection options.


-  **[AllowDeletingColumns](Excel.Protection.AllowDeletingColumns.md)**
    
-  **[AllowDeletingRows](Excel.Protection.AllowDeletingRows.md)**
    
-  **[AllowFiltering](Excel.Protection.AllowFiltering.md)**
    
-  **[AllowFormattingCells](Excel.Protection.AllowFormattingCells.md)**
    
-  **[AllowFormattingColumns](Excel.Protection.AllowFormattingColumns.md)**
    
-  **[AllowFormattingRows](Excel.Protection.AllowFormattingRows.md)**
    
-  **[AllowInsertingColumns](Excel.Protection.AllowInsertingColumns.md)**
    
-  **[AllowInsertingHyperlinks](Excel.Protection.AllowInsertingHyperlinks.md)**
    
-  **[AllowInsertingRows](Excel.Protection.AllowInsertingRows.md)**
    
-  **[AllowSorting](Excel.Protection.AllowSorting.md)**
    
-  **[AllowUsingPivotTables](Excel.Protection.AllowUsingPivotTables.md)**
    

## Example

The following example demonstrates how to use the  **[AllowInsertingColumns](Excel.Protection.AllowInsertingColumns.md)** property of the **Protection** object, placing three numbers in the top row and protecting the worksheet. Then this example checks to see if the protection setting for allowing the insertion of columns is **False** and sets it to **True**, if necessary. Finally, it notifies the user to insert a column.


```
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



|**Name**|
|:-----|
|[AllowDeletingColumns](Excel.Protection.AllowDeletingColumns.md)|
|[AllowDeletingRows](Excel.Protection.AllowDeletingRows.md)|
|[AllowEditRanges](protection-alloweditranges-property-excel.md)|
|[AllowFiltering](Excel.Protection.AllowFiltering.md)|
|[AllowFormattingCells](Excel.Protection.AllowFormattingCells.md)|
|[AllowFormattingColumns](Excel.Protection.AllowFormattingColumns.md)|
|[AllowFormattingRows](Excel.Protection.AllowFormattingRows.md)|
|[AllowInsertingColumns](Excel.Protection.AllowInsertingColumns.md)|
|[AllowInsertingHyperlinks](Excel.Protection.AllowInsertingHyperlinks.md)|
|[AllowInsertingRows](Excel.Protection.AllowInsertingRows.md)|
|[AllowSorting](Excel.Protection.AllowSorting.md)|
|[AllowUsingPivotTables](Excel.Protection.AllowUsingPivotTables.md)|

## See also


#### Other resources


[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)
