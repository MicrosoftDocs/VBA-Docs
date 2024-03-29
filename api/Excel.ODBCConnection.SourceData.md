---
title: ODBCConnection.SourceData property (Excel)
keywords: vbaxl10.chm796088
f1_keywords:
- vbaxl10.chm796088
api_name:
- Excel.ODBCConnection.SourceData
ms.assetid: a23a4c9b-9754-116a-38c8-d71d8f458543
ms.date: 05/01/2019
ms.localizationpriority: medium
---


# ODBCConnection.SourceData property (Excel)

Returns the data source for the ODBC connection, as shown in the table. Read/write **Variant**.


## Syntax

_expression_.**SourceData**

_expression_ A variable that represents an **[ODBCConnection](Excel.ODBCConnection.md)** object.


## Remarks

|Data source|Return value|
|:-----|:-----|
|Microsoft Excel list or database|The cell reference, as text.|
|External data source|An array. Each row consists of an SQL connection string with the remaining elements as the query string, broken down into 255-character segments.|
|Multiple consolidation ranges|A two-dimensional array. Each row consists of a reference and its associated page field items.|
|Another PivotTable report|One of the previous three kinds of information.|

This property is not available for OLE DB data sources.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]