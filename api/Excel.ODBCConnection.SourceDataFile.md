---
title: ODBCConnection.SourceDataFile property (Excel)
keywords: vbaxl10.chm796089
f1_keywords:
- vbaxl10.chm796089
api_name:
- Excel.ODBCConnection.SourceDataFile
ms.assetid: f32c0eeb-e8f5-1a9f-13fd-ead4ad96381f
ms.date: 05/01/2019
ms.localizationpriority: medium
---


# ODBCConnection.SourceDataFile property (Excel)

Returns or sets a **String** indicating the source data file for an ODBC connection. Read/write.


## Syntax

_expression_.**SourceDataFile**

_expression_ A variable that represents an **[ODBCConnection](Excel.ODBCConnection.md)** object.


## Remarks

For file-based data sources (for example, Access) the **SourceDataFile** property contains a fully qualified path to the source data file. It's null for server-based data sources (for example, SQL Server). The **SourceDataFile** property is set to **null** if the **[Connection](Excel.ODBCConnection.Connection.md)** property is changed programmatically.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]