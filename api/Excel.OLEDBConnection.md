---
title: OLEDBConnection Object (Excel)
keywords: vbaxl10.chm793072
f1_keywords:
- vbaxl10.chm793072
ms.prod: excel
api_name:
- Excel.OLEDBConnection
ms.assetid: f246e544-9854-8e71-a7f7-dec57dd725e4
ms.date: 06/08/2017
---


# OLEDBConnection Object (Excel)

Represents the OLE DB connection.


## Remarks

An OLE DB connection can be stored in an Excel workbook. When Micrososft Excel opens the workbook, Excel creates an in-memory copy of the OLE DB connection known as the  **OLEDBConnection** object.

An  **OLEDBConnection** object contains information related to the connection, such as the name of the server to connect to and the name of the objects to be opened on that server. Optionally, the **OLEDBConnection** object may also include authentication credential information, or a command that is to be passed to the server and executed (for example, a `SELECT` statement to be executed by SQL Server).


## Methods



|**Name**|
|:-----|
|[CancelRefresh](Excel.OLEDBConnection.CancelRefresh.md)|
|[MakeConnection](Excel.OLEDBConnection.MakeConnection.md)|
|[Reconnect](Excel.OLEDBConnection.Reconnect.md)|
|[Refresh](Excel.OLEDBConnection.Refresh.md)|
|[SaveAsODC](Excel.OLEDBConnection.SaveAsODC.md)|

## Properties



|**Name**|
|:-----|
|[ADOConnection](Excel.OLEDBConnection.ADOConnection.md)|
|[AlwaysUseConnectionFile](Excel.OLEDBConnection.AlwaysUseConnectionFile.md)|
|[Application](Excel.OLEDBConnection.Application.md)|
|[BackgroundQuery](Excel.OLEDBConnection.BackgroundQuery.md)|
|[CalculatedMembers](Excel.OLEDBConnection.CalculatedMembers.md)|
|[CommandText](Excel.OLEDBConnection.CommandText.md)|
|[CommandType](Excel.OLEDBConnection.CommandType.md)|
|[Connection](Excel.OLEDBConnection.Connection.md)|
|[Creator](Excel.OLEDBConnection.Creator.md)|
|[EnableRefresh](Excel.OLEDBConnection.EnableRefresh.md)|
|[IsConnected](Excel.OLEDBConnection.IsConnected.md)|
|[LocalConnection](Excel.OLEDBConnection.LocalConnection.md)|
|[LocaleID](Excel.OLEDBConnection.LocaleID.md)|
|[MaintainConnection](Excel.OLEDBConnection.MaintainConnection.md)|
|[MaxDrillthroughRecords](Excel.OLEDBConnection.MaxDrillthroughRecords.md)|
|[OLAP](Excel.OLEDBConnection.OLAP.md)|
|[Parent](Excel.OLEDBConnection.Parent.md)|
|[RefreshDate](Excel.OLEDBConnection.RefreshDate.md)|
|[Refreshing](Excel.OLEDBConnection.Refreshing.md)|
|[RefreshOnFileOpen](Excel.OLEDBConnection.RefreshOnFileOpen.md)|
|[RefreshPeriod](Excel.OLEDBConnection.RefreshPeriod.md)|
|[RetrieveInOfficeUILang](Excel.OLEDBConnection.RetrieveInOfficeUILang.md)|
|[RobustConnect](Excel.OLEDBConnection.RobustConnect.md)|
|[SavePassword](Excel.OLEDBConnection.SavePassword.md)|
|[ServerCredentialsMethod](Excel.OLEDBConnection.ServerCredentialsMethod.md)|
|[ServerFillColor](Excel.OLEDBConnection.ServerFillColor.md)|
|[ServerFontStyle](Excel.OLEDBConnection.ServerFontStyle.md)|
|[ServerNumberFormat](Excel.OLEDBConnection.ServerNumberFormat.md)|
|[ServerSSOApplicationID](Excel.OLEDBConnection.ServerSSOApplicationID.md)|
|[ServerTextColor](Excel.OLEDBConnection.ServerTextColor.md)|
|[SourceConnectionFile](Excel.OLEDBConnection.SourceConnectionFile.md)|
|[SourceDataFile](Excel.OLEDBConnection.SourceDataFile.md)|
|[UseLocalConnection](Excel.OLEDBConnection.UseLocalConnection.md)|

## See also


#### Other resources


[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)
