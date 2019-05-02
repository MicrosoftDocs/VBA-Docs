---
title: PivotCache.CommandType property (Excel)
keywords: vbaxl10.chm227088
f1_keywords:
- vbaxl10.chm227088
ms.prod: excel
api_name:
- Excel.PivotCache.CommandType
ms.assetid: bbe0ba26-efb9-428d-de2c-576116d92747
ms.date: 05/03/2019
localization_priority: Normal
---


# PivotCache.CommandType property (Excel)

Returns or sets one of these **[XlCmdType](Excel.XlCmdType.md)** constants: **xlCmdCube**, **xlCmdDefault**, **xlCmdSql**, or **xlCmdTable**. 

The constant that is returned or set describes the value of the **[CommandText](Excel.PivotCache.CommandText.md)** property. The default value is **xlCmdSQL**. Read/write **XlCmdType**.


## Syntax

_expression_.**CommandType**

_expression_ An expression that returns a **[PivotCache](Excel.PivotCache.md)** object.


## Remarks

You can set the **CommandType** property only if the value of the **[QueryType](Excel.PivotCache.QueryType.md)** property for the query table or PivotTable cache is **xlOLEDBQuery**.

If the value of the **CommandType** property is **xlCmdCube**, you cannot change this value if there is a PivotTable report associated with the query table.


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