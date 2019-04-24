---
title: QueryTable.QueryType property (Excel)
keywords: vbaxl10.chm518116
f1_keywords:
- vbaxl10.chm518116
ms.prod: excel
api_name:
- Excel.QueryTable.QueryType
ms.assetid: 7cf9ea40-62ea-7211-7832-31eceb44ed15
ms.date: 06/08/2017
localization_priority: Normal
---


# QueryTable.QueryType property (Excel)

Indicates the type of query used by Microsoft Excel to populate the query table. Read-only  **[XlQueryType](Excel.XlQueryType.md)**.


## Syntax

_expression_. `QueryType`

_expression_ A variable that represents a [QueryTable](Excel.QueryTable.md) object.


## Remarks



| **xlQueryType** can be one of these **xlQueryType** constants.|
| **xlTextImport**. Based on a text file, for query tables only|
| **xlOLEDBQuery**. Based on an OLE DB query, including OLAP data sources|
| **xlWebQuery**. Based on a webpage, for query tables only|
| **xlADORecordset**. Based on an ADO recordset query|
| **xlDAORecordSet**. Based on a DAO recordset query, for query tables only|
| **xlODBCQuery**. Based on an ODBC data source|

You specify the data source in the prefix for the  **[Connection](Excel.QueryTable.Connection.md)** property's value.

If you import data using the user interface, data from a web query or a text query is imported as a  **[QueryTable](Excel.QueryTable.md)** object, while all other external data is imported as a **[ListObject](Excel.ListObject.md)** object.

If you import data using the object model, data from a web query or a text query must be imported as a  **QueryTable**, while all other external data can be imported as either a **ListObject** or a **QueryTable**.

You can use the  **[QueryTable](Excel.ListObject.QueryTable.md)** property of the **ListObject** to access the **QueryType** property.


## Example

This example refreshes the first query table on the first worksheet if the table is based on a webpage.


```vb
Set qtQtrResults = _ 
 Workbooks(1).Worksheets(1).QueryTables(1) 
With qtQtrResults 
 if .QueryType = xlWebQuery Then 
 .Refresh 
 End If 
End With
```


## See also


[QueryTable Object](Excel.QueryTable.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]