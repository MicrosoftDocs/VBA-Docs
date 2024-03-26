---
title: PivotCache.CreatePivotTable method (Excel)
keywords: vbaxl10.chm227095
f1_keywords:
- vbaxl10.chm227095
api_name:
- Excel.PivotCache.CreatePivotTable
ms.assetid: dca20930-5d58-8db7-bd81-3c90b7588011
ms.date: 05/03/2019
ms.localizationpriority: medium
---


# PivotCache.CreatePivotTable method (Excel)

Creates a PivotTable report based on a **PivotCache** object. Returns a **[PivotTable](Excel.PivotTable.md)** object.


## Syntax

_expression_.**CreatePivotTable** (_TableDestination_, _TableName_, _ReadData_, _DefaultVersion_)

_expression_ A variable that represents a **[PivotCache](Excel.PivotCache.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _TableDestination_|Required| **Variant**|The cell in the upper-left corner of the PivotTable report's destination range (the range on the worksheet where the resulting PivotTable report will be placed). The destination range must be on a worksheet in the workbook that contains the **PivotCache** object specified by _expression_.|
| _TableName_|Optional| **Variant**|The name of the new PivotTable report.|
| _ReadData_|Optional| **Variant**| **True** to create a PivotTable cache that contains all the records from the external database; this cache can be very large. **False** to enable setting some of the fields as server-based page fields before the data is actually read.|
| _DefaultVersion_|Optional| **Variant**|The default version of the PivotTable report.|

## Return value

PivotTable


## Remarks

For an alternative way to create a PivotTable report based on a PivotTable cache, see the **[Add](Excel.PivotTables.Add.md)** method of the **PivotTables** object.


## Example

This example creates a new PivotTable cache based on an OLAP provider, and then it creates a new PivotTable report based on the cache at cell A3 on the active worksheet.

```vb
With ActiveWorkbook.PivotCaches.Add(SourceType:=xlExternal) 
 .Connection = _ 
 "OLEDB;Provider=MSOLAP;Location=srvdata;Initial Catalog=National" 
 .CommandType = xlCmdCube 
 .CommandText = Array("Sales") 
 .MaintainConnection = True 
 .CreatePivotTable TableDestination:=Range("A3"), _ 
 TableName:= "PivotTable1" 
End With 
With ActiveSheet.PivotTables("PivotTable1") 
 .SmallGrid = False 
 .PivotCache.RefreshPeriod = 0 
 With .CubeFields("[state]") 
 .Orientation = xlColumnField 
 .Position = 1 
 End With 
 With .CubeFields("[Measures].[Count Of au_id]") 
 .Orientation = xlDataField 
 .Position = 1 
 End With 
End With
```

This example creates a new PivotTable cache by using an ADO connection to Microsoft Jet, and then it creates a new PivotTable report based on the cache at cell A3 on the active worksheet.

```vb
Dim cnnConn As ADODB.Connection 
Dim rstRecordset As ADODB.Recordset 
Dim cmdCommand As ADODB.Command 
 
' Open the connection. 
Set cnnConn = New ADODB.Connection 
With cnnConn 
 .ConnectionString = _ 
 "Provider=Microsoft.Jet.OLEDB.4.0" 
 .Open "C:\perfdate\record.mdb" 
End With 
 
' Set the command text. 
Set cmdCommand = New ADODB.Command 
Set cmdCommand.ActiveConnection = cnnConn 
With cmdCommand 
 .CommandText = "Select Speed, Pressure, Time From DynoRun" 
 .CommandType = adCmdText 
 .Execute 
End With 
 
' Open the recordset. 
Set rstRecordset = New ADODB.Recordset 
Set rstRecordset.ActiveConnection = cnnConn 
rstRecordset.Open cmdCommand 
 
' Create a PivotTable cache and report. 
Set objPivotCache = ActiveWorkbook.PivotCaches.Add( _ 
 SourceType:=xlExternal) 
Set objPivotCache.Recordset = rstRecordset 
With objPivotCache 
 .CreatePivotTable TableDestination:=Range("A3"), _ 
 TableName:="Performance" 
End With 
 
With ActiveSheet.PivotTables("Performance") 
 .SmallGrid = False 
 With .PivotFields("Pressure") 
 .Orientation = xlRowField 
 .Position = 1 
 End With 
 With .PivotFields("Speed") 
 .Orientation = xlColumnField 
 .Position = 1 
 End With 
 With .PivotFields("Time") 
 .Orientation = xlDataField 
 .Position = 1 
 End With 
End With 
 
' Close the connections and clean up. 
cnnConn.Close 
Set cmdCommand = Nothing 
Set rstRecordSet = Nothing 
Set cnnConn = Nothing
```
This example use already existing WorkbookConnection.
```vb
 
 'Get WorkbookConnection object
 Dim conn As WorkbookConnection
 Set conn =  ActiveWorkbook.Connections("MyConnectionName")
 
 'Declare temp variables
 Dim connStr As String
 Dim sqlStr As String
 
 'Store connection string and command text in variables depends on connection type
 If conn.Type = xlConnectionTypeODBC Then
   connStr = conn.ODBCConnection.Connection
   sqlStr = conn.ODBCConnection.CommandText
 End If
  
 If conn.Type = xlConnectionTypeOLEDB Then
   connStr = conn.OLEDBConnection.Connection
   sqlStr = conn.OLEDBConnection.CommandText
 End If
 
 'Create PivotCache
 Dim pcache As pivotCache
 Set pcache = ActiveWorkbook.PivotCaches.Create(xlExternal, conn)
 
 'Then we need to get recordset to create pivot table
 Dim adodb_conn As Object
 Dim rs As Object
 Set adodb_conn = CreateObject("ADODB.Connection")
 Set rs = CreateObject("ADODB.Recordset")
 adodb_conn.Open connStr
 rs.Open sqlStr, adodb_conn
 
 Set pcache.Recordset = rs
 'When CreatePivotTable method called the linked WorkbookConnection is losing connection string and command text
 Set pvt = pcache.CreatePivotTable(TableDestination:=Sheets("MySheetName").Cells(1, 1), TableName:="MyPivotTableName")
        
 rs.Close
 adodb_conn.Close
 
 'Restore CommandText and connection string
 pcache.CommandText = sqlStr
 pcache.Connection = connStr
 
 ' Now you have PivotTable that linked with yours WorkbookConnection
 
 
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
