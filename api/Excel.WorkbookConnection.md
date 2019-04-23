---
title: WorkbookConnection object (Excel)
keywords: vbaxl10.chm773072
f1_keywords:
- vbaxl10.chm773072
ms.prod: excel
api_name:
- Excel.WorkbookConnection
ms.assetid: 5974dd57-7671-cd55-3f8f-6a76fa938317
ms.date: 04/03/2019
localization_priority: Normal
---


# WorkbookConnection object (Excel)

A connection is a set of information needed to obtain data from an external data source other than a Microsoft Excel workbook. 


## Remarks

### Store connections in an Excel workbook

Connections can be stored within an Excel workbook. When the workbook is opened, Excel creates an in-memory copy of the connection that is referred to as the connection object. A connection object contains information such as the name of the server and the name of the object to be opened on that server. 

Optionally, the connection object may also include authentication credentials and/or a command that is to be passed to the server and executed (example: a SELECT statement to be executed by SQL Server).

> [!NOTE] 
> The exact form of the connection depends on the mechanism that is being used to retrieve data; ODBC connections, OLEDB connections, and web queries will contain different information.

### Store connections in a connection file

A connection may also be stored in a separate connection file. Most connections in an Excel workbook include a pointer to an external connection file. Connection files have extensions that clearly label them as connection files (*.ODC, *.IQY, etc.) and may be located on the user's local machine or in other well-known or trusted locations such as WSS (Data Connection Library), or other corporate servers. 

Connection files enable multiple users within the same organization to re-use connections. Network administrators are able to change the way the entire organization connects to a back-end data source by changing a single connection file. A connection file is not always required when connecting to an external data source.

### Identify connections

Connection names are strings that uniquely identify connections within the workbook in which they are used. There are other properties of a connection that are not unique. Whenever a formula in Excel takes an argument that is a connection, it will be sufficient to refer to the name of that connection, either directly (as a string) or indirectly (by referring to a cell that contains the connection name as a string).


## Methods

- [Delete](Excel.WorkbookConnection.Delete.md)
- [Refresh](Excel.WorkbookConnection.Refresh.md)

## Properties

- [Application](Excel.WorkbookConnection.Application.md)
- [Creator](Excel.WorkbookConnection.Creator.md)
- [DataFeedConnection](Excel.workbookconnection.datafeedconnection.md)
- [Description](Excel.WorkbookConnection.Description.md)
- [InModel](Excel.workbookconnection.inmodel.md)
- [ModelConnection](Excel.workbookconnection.modelconnection.md)
- [ModelTables](Excel.workbookconnection.modeltables.md)
- [Name](Excel.WorkbookConnection.Name.md)
- [ODBCConnection](Excel.WorkbookConnection.ODBCConnection.md)
- [OLEDBConnection](Excel.WorkbookConnection.OLEDBConnection.md)
- [Parent](Excel.WorkbookConnection.Parent.md)
- [Ranges](Excel.WorkbookConnection.Ranges.md)
- [RefreshWithRefreshAll](Excel.workbookconnection.refreshwithrefreshall.md)
- [TextConnection](Excel.workbookconnection.textconnection.md)
- [Type](Excel.WorkbookConnection.Type.md)
- [WorksheetDataConnection](Excel.workbookconnection.worksheetdataconnection.md)


## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
