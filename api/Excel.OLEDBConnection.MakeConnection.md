---
title: OLEDBConnection.MakeConnection method (Excel)
keywords: vbaxl10.chm794082
f1_keywords:
- vbaxl10.chm794082
ms.prod: excel
api_name:
- Excel.OLEDBConnection.MakeConnection
ms.assetid: ff618eae-1593-aabc-dbcb-427291caf923
ms.date: 05/02/2019
localization_priority: Normal
---


# OLEDBConnection.MakeConnection method (Excel)

Establishes a connection for the specified OLE DB connection.


## Syntax

_expression_.**MakeConnection**

_expression_ A variable that represents an **[OLEDBConnection](Excel.OLEDBConnection.md)** object.


## Return value

Nothing


## Remarks

The **MakeConnection** method can be used when a connection drops and the user wants to reestablish the connection.

Various objects and methods might return a run-time error if the connection is dropped. Use of this method assures a connection before executing other objects or methods.

> [!NOTE] 
> Microsoft Excel might drop a connection temporarily in the course of a session (unknown to the VBA programmer), so this method proves useful.

This method will result in a run-time error if the **[MaintainConnection](Excel.OLEDBConnection.MaintainConnection.md)** property of the specified OLE DB connection has been set to **False**.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]