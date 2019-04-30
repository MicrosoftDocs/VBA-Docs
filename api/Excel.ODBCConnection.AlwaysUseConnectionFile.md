---
title: ODBCConnection.AlwaysUseConnectionFile property (Excel)
keywords: vbaxl10.chm796092
f1_keywords:
- vbaxl10.chm796092
ms.prod: excel
api_name:
- Excel.ODBCConnection.AlwaysUseConnectionFile
ms.assetid: 445c7371-0ac6-b6f3-1a78-a406922d106f
ms.date: 05/01/2019
localization_priority: Normal
---


# ODBCConnection.AlwaysUseConnectionFile property (Excel)

**True** if the connection file is always used to establish connection to the data source. Read/write **Boolean**.


## Syntax

_expression_.**AlwaysUseConnectionFile**

_expression_ A variable that represents an **[ODBCConnection](Excel.ODBCConnection.md)** object.


## Remarks

When this property is **True**, the connection file will be used to establish the connection to the data source. If the connection embedded within the workbook is different from the external connection file, the embedded connection will be ignored and the external connection file will be the only version considered.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]