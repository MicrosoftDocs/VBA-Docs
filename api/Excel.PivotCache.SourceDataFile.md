---
title: PivotCache.SourceDataFile property (Excel)
keywords: vbaxl10.chm227104
f1_keywords:
- vbaxl10.chm227104
ms.prod: excel
api_name:
- Excel.PivotCache.SourceDataFile
ms.assetid: 1b90ee17-45c1-3c96-33e3-ec6c5515d9ee
ms.date: 05/03/2019
localization_priority: Normal
---


# PivotCache.SourceDataFile property (Excel)

Returns a **String** value that indicates the source data file for the cache of the PivotTable.


## Syntax

_expression_.**SourceDataFile**

_expression_ A variable that represents a **[PivotCache](Excel.PivotCache.md)** object.


## Remarks

For file-based data sources (for example, Access), the **SourceDataFile** property contains a fully qualified path to the source data file. It is set to **Null** for server-based data sources (such as SQL Server). The **SourceDataFile** property is set to **Null** if the **[Connection](Excel.PivotCache.Connection.md)** property is changed programmatically.


## Example

This example determines if a connection exists for the cache, and if there is a connection, displays the data source file name. If no connection exists, the code handles the run-time error and notifies the user. This example assumes that a PivotTable exists on the active worksheet.

```vb
Sub CheckSourceConnection() 
 
 Dim pvtCache As PivotCache 
 
 Set pvtCache = Application.ActiveWorkbook.PivotCaches.Item(1) 
 
 On Error GoTo No_Connection 
 
 MsgBox "The data source connection is: " & _ 
 pvtCache.SourceDataFile 
 Exit Sub 
 
No_Connection: 
 MsgBox "PivotCache source cannot be determined." 
 
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]