---
title: ODBCConnection.RobustConnect property (Excel)
keywords: vbaxl10.chm796084
f1_keywords:
- vbaxl10.chm796084
ms.prod: excel
api_name:
- Excel.ODBCConnection.RobustConnect
ms.assetid: 2f575278-d385-90bd-6544-885f99abbebb
ms.date: 05/01/2019
localization_priority: Normal
---


# ODBCConnection.RobustConnect property (Excel)

Returns or sets how an ODBC connection connects to its data source. Read/write **[XlRobustConnect](Excel.XlRobustConnect.md)**.


## Syntax

_expression_.**RobustConnect**

_expression_ A variable that represents an **[ODBCConnection](Excel.ODBCConnection.md)** object.


## Remarks

If you use robust connectivity, when Microsoft Excel is unable to establish a connection by using the workbook connection information, Excel will check if the workbook connection has a path to an external connection file. If it does, Excel will open the external connection file and try to connect by using the information contained in the external connection file. If a connection can be established after using the external connection file, Excel will then update the workbook's connection object. 

This provides robust connectivity in scenarios where an information technology department maintains and updates connections in a central place, permitting a user's workbook to automatically fetch those updates whenever the previous version of the connection (cached within the workbook) fails. 

> [!NOTE] 
> Robust connectivity isn't secure. If you create a connection on your computer and use it in a workbook and then send someone the workbook, that person will be able to see the path to the connection file on your computer. It is a good idea to remove the connection file information from the workbook before you send the workbook to someone else.



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]