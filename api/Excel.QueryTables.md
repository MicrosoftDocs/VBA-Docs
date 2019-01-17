---
title: QueryTables object (Excel)
keywords: vbaxl10.chm520072
f1_keywords:
- vbaxl10.chm520072
ms.prod: excel
api_name:
- Excel.QueryTables
ms.assetid: 93511da3-598e-0aa3-fbc3-14bebff8838f
ms.date: 06/08/2017
localization_priority: Priority
---


# QueryTables object (Excel)

A collection of  **[QueryTable](Excel.QueryTable.md)** objects.


## Remarks

 Each **QueryTable** object represents a worksheet table built from data returned from an external data source.


## Example

Use the  **[QueryTables](Excel.Worksheet.QueryTables.md)** property to return the **[QueryTables](Excel.QueryTables.md)** collection. The following example displays the number of query tables on the active worksheet.


```vb
MsgBox ActiveSheet.QueryTables.Count
```

Use the  **[Add](Excel.QueryTables.Add.md)** method to create a new query table and add it to the **QueryTables** collection. The following example creates a new query table.




```vb
Dim qt As QueryTable 
sqlstring = "select 96Sales.totals from 96Sales where profit < 5" 
connstring = _ 
 "ODBC;DSN=96SalesData;UID=Rep21;PWD=NUyHwYQI;Database=96Sales" 
With ActiveSheet.QueryTables.Add(Connection:=connstring, _ 
 Destination:=Range("B1"), Sql:=sqlstring) 
 .Refresh 
End With
```


## Methods



|Name|
|:-----|
|[Add](Excel.QueryTables.Add.md)|
|[Item](Excel.QueryTables.Item.md)|

## Properties



|Name|
|:-----|
|[Application](Excel.QueryTables.Application.md)|
|[Count](Excel.QueryTables.Count.md)|
|[Creator](Excel.QueryTables.Creator.md)|
|[Parent](Excel.QueryTables.Parent.md)|

## See also


[Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]