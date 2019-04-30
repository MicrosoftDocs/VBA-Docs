---
title: ODBCConnection.CommandType property (Excel)
keywords: vbaxl10.chm796076
f1_keywords:
- vbaxl10.chm796076
ms.prod: excel
api_name:
- Excel.ODBCConnection.CommandType
ms.assetid: 5bfffa11-94d1-43fa-1da5-83f341c0a3cd
ms.date: 05/01/2019
localization_priority: Normal
---


# ODBCConnection.CommandType property (Excel)

Returns or sets one of the **XlCmdType** constants. Read/write **[XlCmdType](Excel.XlCmdType.md)**.


## Syntax

_expression_.**CommandType**

_expression_ A variable that represents an **[ODBCConnection](Excel.ODBCConnection.md)** object.


## Remarks

The constant that is returned or set describes the value of the **[CommandText](Excel.ODBCConnection.CommandText.md)** property. The default value is **xlCmdSQL**.


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