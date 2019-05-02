---
title: PivotCache.Recordset property (Excel)
keywords: vbaxl10.chm227092
f1_keywords:
- vbaxl10.chm227092
ms.prod: excel
api_name:
- Excel.PivotCache.Recordset
ms.assetid: 25f2eb4f-d78c-21e2-9d26-c8ebc3404607
ms.date: 05/03/2019
localization_priority: Normal
---


# PivotCache.Recordset property (Excel)

Returns or sets a **Recordset** object that's used as the data source for the specified PivotTable cache. Read/write.


## Syntax

_expression_.**Recordset**

_expression_ A variable that represents a **[PivotCache](Excel.PivotCache.md)** object.


## Remarks

If this property is used to overwrite an existing recordset, the change takes effect when the **[Refresh](Excel.PivotCache.Refresh.md)** method is run.


## Example

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




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]