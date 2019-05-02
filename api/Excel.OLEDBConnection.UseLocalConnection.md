---
title: OLEDBConnection.UseLocalConnection property (Excel)
keywords: vbaxl10.chm794094
f1_keywords:
- vbaxl10.chm794094
ms.prod: excel
api_name:
- Excel.OLEDBConnection.UseLocalConnection
ms.assetid: b346933c-17cd-ef11-6070-ee840c8d7c0a
ms.date: 05/02/2019
localization_priority: Normal
---


# OLEDBConnection.UseLocalConnection property (Excel)

**True** if the **[LocalConnection](Excel.OLEDBConnection.LocalConnection.md)** property is used to specify the string that enables Microsoft Excel to connect to a data source. **False** if the connection string specified by the **[Connection](Excel.OLEDBConnection.Connection.md)** property is used. Read/write **Boolean**.


## Syntax

_expression_.**UseLocalConnection**

_expression_ A variable that represents an **[OLEDBConnection](Excel.OLEDBConnection.md)** object.


## Example

This example sets the connection string of the first PivotTable cache to reference an offline cube file.

```vb
With ActiveWorkbook.PivotCaches(1) 
 .LocalConnection = _ 
 "OLEDB;Provider=MSOLAP;Data Source=C:\Data\DataCube.cub" 
 .UseLocalConnection = True 
End With 

```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]