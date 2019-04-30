---
title: ODBCConnection.Connection property (Excel)
keywords: vbaxl10.chm796077
f1_keywords:
- vbaxl10.chm796077
ms.prod: excel
api_name:
- Excel.ODBCConnection.Connection
ms.assetid: 2fcd1043-b088-cfde-9853-4a20da20be26
ms.date: 05/01/2019
localization_priority: Normal
---


# ODBCConnection.Connection property (Excel)

Returns or sets a string that contains ODBC settings that enable Microsoft Excel to connect to an ODBC data source. Read/write **Variant**.


## Syntax

_expression_.**Connection**

_expression_ A variable that represents an **[ODBCConnection](Excel.ODBCConnection.md)** object.


## Remarks

Setting the **Connection** property does not immediately initiate the connection to the data source. You must use the **[Refresh](Excel.ODBCConnection.Refresh.md)** method to make the connection and retrieve the data.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]