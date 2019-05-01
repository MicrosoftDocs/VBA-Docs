---
title: OLEDBConnection.IsConnected property (Excel)
keywords: vbaxl10.chm794096
f1_keywords:
- vbaxl10.chm794096
ms.prod: excel
api_name:
- Excel.OLEDBConnection.IsConnected
ms.assetid: 3538c8bd-5027-8f48-d6b5-b18de0db4159
ms.date: 05/02/2019
localization_priority: Normal
---


# OLEDBConnection.IsConnected property (Excel)

Returns **True** if the **[MaintainConnection](Excel.OLEDBConnection.MaintainConnection.md)** property is **True**. Returns **False** if it is not currently connected to its source. Read-only **Boolean**.


## Syntax

_expression_.**IsConnected**

_expression_ A variable that represents an **[OLEDBConnection](Excel.OLEDBConnection.md)** object.


## Remarks

The **IsConnected** property does not check to see if the connection is connected. Even if this property returns **True**, sending commands to the provider could result in an error if the connection is no longer valid.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]