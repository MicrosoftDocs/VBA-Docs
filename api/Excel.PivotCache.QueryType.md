---
title: PivotCache.QueryType property (Excel)
keywords: vbaxl10.chm227089
f1_keywords:
- vbaxl10.chm227089
ms.prod: excel
api_name:
- Excel.PivotCache.QueryType
ms.assetid: 61346ed2-1ada-a105-1894-b22861047c4f
ms.date: 06/08/2017
localization_priority: Normal
---


# PivotCache.QueryType property (Excel)

Indicates the type of query used by Microsoft Excel to populate the PivotTable cache. Read-only  **[XlQueryType](Excel.XlQueryType.md)**.


## Syntax

_expression_. `QueryType`

_expression_ A variable that represents a **[PivotCache](Excel.PivotCache.md)** object.


## Remarks



| **xlQueryType** can be one of these **xlQueryType** constants.|
| **xlTextImport**. Based on a text file, for query tables only|
| **xlOLEDBQuery**. Based on an OLE DB query, including OLAP data sources|
| **xlWebQuery**. Based on a webpage, for query tables only|
| **xlADORecordset**. Based on an ADO recordset query|
| **xlDAORecordSet**. Based on a DAO recordset query, for query tables only|
| **xlODBCQuery**. Based on an ODBC data source|

You specify the data source in the prefix for the  **[Connection](Excel.PivotCache.Connection.md)** property's value.


## See also


[PivotCache Object](Excel.PivotCache.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]