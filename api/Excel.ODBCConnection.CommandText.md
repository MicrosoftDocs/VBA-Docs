---
title: ODBCConnection.CommandText property (Excel)
keywords: vbaxl10.chm796075
f1_keywords:
- vbaxl10.chm796075
ms.prod: excel
api_name:
- Excel.ODBCConnection.CommandText
ms.assetid: f76073fd-5052-5813-ee9a-631c795e9b76
ms.date: 05/01/2019
localization_priority: Normal
---


# ODBCConnection.CommandText property (Excel)

Returns or sets the command string for the specified data source. Read/write **Variant**.


## Syntax

_expression_.**CommandText**

_expression_ A variable that represents an **[ODBCConnection](Excel.ODBCConnection.md)** object.


## Remarks

You should use the **CommandText** property instead of the **SQL** property, which now exists primarily for compatibility with earlier versions of Microsoft Excel. If you use both properties, the **CommandText** property's value takes precedence.

The **[CommandType](Excel.ODBCConnection.CommandType.md)** property describes the value of the **CommandText** property.


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