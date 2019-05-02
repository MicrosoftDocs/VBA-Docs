---
title: QueryTable.QueryType property (Excel)
keywords: vbaxl10.chm518116
f1_keywords:
- vbaxl10.chm518116
ms.prod: excel
api_name:
- Excel.QueryTable.QueryType
ms.assetid: 7cf9ea40-62ea-7211-7832-31eceb44ed15
ms.date: 05/03/2019
localization_priority: Normal
---


# QueryTable.QueryType property (Excel)

Indicates the type of query used by Microsoft Excel to populate the query table. Read-only **[XlQueryType](Excel.XlQueryType.md)**.


## Syntax

_expression_.**QueryType**

_expression_ A variable that represents a **[QueryTable](Excel.QueryTable.md)** object.


## Remarks

You specify the data source in the prefix for the **[Connection](Excel.QueryTable.Connection.md)** property's value.

If you import data by using the user interface, data from a web query or a text query is imported as a **QueryTable** object, while all other external data is imported as a **[ListObject](Excel.ListObject.md)** object.

If you import data by using the object model, data from a web query or a text query must be imported as a **QueryTable**, while all other external data can be imported as either a **ListObject** or a **QueryTable**.

You can use the **[QueryTable](Excel.ListObject.QueryTable.md)** property of the **ListObject** to access the **QueryType** property.


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




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]