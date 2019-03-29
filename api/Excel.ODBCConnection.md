---
title: ODBCConnection object (Excel)
keywords: vbaxl10.chm795072
f1_keywords:
- vbaxl10.chm795072
ms.prod: excel
api_name:
- Excel.ODBCConnection
ms.assetid: b880ebec-15a4-5a3d-ef02-db73106db9c9
ms.date: 03/30/2019
localization_priority: Normal
---


# ODBCConnection object (Excel)

Represents the ODBC connection.


## Remarks

An ODBC connection can be stored in an Excel workbook. When Excel opens the workbook, it creates an in-memory copy of the ODBC connection known as the **ODBCConnection** object.

An **ODBCConnection** object contains information related to the connection, such as the name of the server to connect to and the name of the objects to be opened on that server. Optionally, the **ODBCConnection** object may also include authentication credential information, or a command that is to be passed to the server and executed (for example, a **SELECT** statement to be executed by SQL Server).

## Methods

- [CancelRefresh](Excel.ODBCConnection.CancelRefresh.md)
- [Refresh](Excel.ODBCConnection.Refresh.md)
- [SaveAsODC](Excel.ODBCConnection.SaveAsODC.md)

## Properties

- [AlwaysUseConnectionFile](Excel.ODBCConnection.AlwaysUseConnectionFile.md)
- [Application](Excel.ODBCConnection.Application.md)
- [BackgroundQuery](Excel.ODBCConnection.BackgroundQuery.md)
- [CommandText](Excel.ODBCConnection.CommandText.md)
- [CommandType](Excel.ODBCConnection.CommandType.md)
- [Connection](Excel.ODBCConnection.Connection.md)
- [Creator](Excel.ODBCConnection.Creator.md)
- [EnableRefresh](Excel.ODBCConnection.EnableRefresh.md)
- [Parent](Excel.ODBCConnection.Parent.md)
- [RefreshDate](Excel.ODBCConnection.RefreshDate.md)
- [Refreshing](Excel.ODBCConnection.Refreshing.md)
- [RefreshOnFileOpen](Excel.ODBCConnection.RefreshOnFileOpen.md)
- [RefreshPeriod](Excel.ODBCConnection.RefreshPeriod.md)
- [RobustConnect](Excel.ODBCConnection.RobustConnect.md)
- [SavePassword](Excel.ODBCConnection.SavePassword.md)
- [ServerCredentialsMethod](Excel.ODBCConnection.ServerCredentialsMethod.md)
- [ServerSSOApplicationID](Excel.ODBCConnection.ServerSSOApplicationID.md)
- [SourceConnectionFile](Excel.ODBCConnection.SourceConnectionFile.md)
- [SourceData](Excel.ODBCConnection.SourceData.md)
- [SourceDataFile](Excel.ODBCConnection.SourceDataFile.md)


## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]