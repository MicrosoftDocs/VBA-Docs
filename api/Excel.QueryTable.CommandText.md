---
title: QueryTable.CommandText property (Excel)
keywords: vbaxl10.chm518113
f1_keywords:
- vbaxl10.chm518113
ms.prod: excel
api_name:
- Excel.QueryTable.CommandText
ms.assetid: 5f1f84f2-d613-17be-7b2e-3b6a3cc56002
ms.date: 05/03/2019
localization_priority: Normal
---


# QueryTable.CommandText property (Excel)

Returns or sets the command string for the specified data source. Read/write **Variant**.


## Syntax

_expression_.**CommandText**

_expression_ An expression that returns a **[QueryTable](Excel.QueryTable.md)** object.


## Remarks

For OLE DB sources, the **[CommandType](Excel.QueryTable.CommandType.md)** property describes the value of the **CommandText** property.

For ODBC sources, setting the **CommandText** causes the data to be refreshed.

If you import data by using the user interface, data from a web query or a text query is imported as a **QueryTable** object, while all other external data is imported as a **[ListObject](Excel.ListObject.md)** object.

If you import data by using the object model, data from a web query or a text query must be imported as a **QueryTable**, while all other external data can be imported as either a **ListObject** or a **QueryTable**.

You can use the **[QueryTable](Excel.ListObject.QueryTable.md)** property of the **ListObject** to access the **CommandText** property.

The sheet that contains the query table must be active to access this property.


## Example

This example sets the command string for the first query table's ODBC data source. Note that the command string is an SQL statement.

```vb
Set qtQtrResults = _ 
 Workbooks(1).Worksheets(1).QueryTables(1) 
With qtQtrResults 
 .CommandType = xlCmdSQL 
 .CommandText = _ 
 "Select ProductID From Products Where ProductID < 10" 
 .Refresh 
End With
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
