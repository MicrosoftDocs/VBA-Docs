---
title: QueryTable.RefreshStyle property (Excel)
keywords: vbaxl10.chm518083
f1_keywords:
- vbaxl10.chm518083
ms.prod: excel
api_name:
- Excel.QueryTable.RefreshStyle
ms.assetid: d32e96f9-ab4f-c6d5-50ac-13c9b1939a0f
ms.date: 06/08/2017
localization_priority: Priority
---


# QueryTable.RefreshStyle property (Excel)

Returns or sets the way rows on the specified worksheet are added or deleted to accommodate the number of rows in a recordset returned by a query. Read/write  **[xlCellInsertionMode](Excel.XlCellInsertionMode.md)**.


## Syntax

_expression_. `RefreshStyle`

_expression_ A variable that represents a [QueryTable](Excel.QueryTable.md) object.


## Remarks



| **xlCellInsertionMode** can be one of these **xlCellInsertionMode** constants.|
| **xlInsertDeleteCells** Partial rows are inserted or deleted to match the exact number of rows required for the new recordset.|
| **xlOverwriteCells** No new cells or rows are added to the worksheet. Data in surrounding cells is overwritten to accommodate any overflow.|
| **xlInsertEntireRows** Entire rows are inserted, if necessary, to accommodate any overflow. No cells or rows are deleted from the worksheet.|

If you import data using the user interface, data from a Web query or a text query is imported as a  **[QueryTable](Excel.QueryTable.md)** object, while all other external data is imported as a **[ListObject](Excel.ListObject.md)** object.

If you import data using the object model, data from a Web query or a text query must be imported as a  **QueryTable** , while all other external data can be imported as either a **ListObject** or a **QueryTable**.

You can use the  **[QueryTable](Excel.ListObject.QueryTable.md)** property of the **ListObject** to access the **RefreshStyle** property.


## Example

This example adds a query table to Sheet1. The  **RefreshStyle** property adds rows to the worksheet as needed, to hold the data results.


```vb
Dim qt As QueryTable 
Set qt = Sheets("sheet1").QueryTables _ 
 .Add(Connection:="Finder;c:\myfile.dqy", _ 
 Destination:=Range("sheet1!a1")) 
With qt 
 .RefreshStyle = xlInsertEntireRows 
 .Refresh 
End With
```


## See also


[QueryTable Object](Excel.QueryTable.md)

