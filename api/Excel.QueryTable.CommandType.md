---
title: QueryTable.CommandType property (Excel)
keywords: vbaxl10.chm518114
f1_keywords:
- vbaxl10.chm518114
ms.prod: excel
api_name:
- Excel.QueryTable.CommandType
ms.assetid: ed1b668c-a73c-0ee7-45ed-67a9d46921dd
ms.date: 05/03/2019
localization_priority: Normal
---


# QueryTable.CommandType property (Excel)

Returns or sets one of these **[XlCmdType](Excel.XlCmdType.md)** constants: **xlCmdCube**, **xlCmdDefault**, **xlCmdSql**, or **xlCmdTable**. The constant that is returned or set describes the value of the **[CommandText](Excel.QueryTable.CommandText.md)** property. The default value is **xlCmdSQL**. Read/write **XlCmdType**.


## Syntax

_expression_.**CommandType**

_expression_ An expression that returns a **[QueryTable](Excel.QueryTable.md)** object.


## Remarks

You can set the **CommandType** property only if the value of the **[QueryType](Excel.QueryTable.QueryType.md)** property for the query table or PivotTable cache is **xlOLEDBQuery**.

If the value of the **CommandType** property is **xlCmdCube**, you cannot change this value if there is a PivotTable report associated with the query table.

If you import data by using the user interface, data from a web query or a text query is imported as a **QueryTable** object, while all other external data is imported as a **[ListObject](Excel.ListObject.md)** object.

If you import data by using the object model, data from a web query or a text query must be imported as a **QueryTable**, while all other external data can be imported as either a **ListObject** or a **QueryTable**.

You can use the **[QueryTable](Excel.ListObject.QueryTable.md)** property of the **ListObject** to access the **CommandType** property.


## Example

This example sets the command string for the first query table's ODBC data source. The command string is an SQL statement.

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
