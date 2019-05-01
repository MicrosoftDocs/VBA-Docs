---
title: OLEDBConnection.AlwaysUseConnectionFile property (Excel)
keywords: vbaxl10.chm794099
f1_keywords:
- vbaxl10.chm794099
ms.prod: excel
api_name:
- Excel.OLEDBConnection.AlwaysUseConnectionFile
ms.assetid: de9cd9a7-0dd6-7ee2-d48f-bd61a7006c1e
ms.date: 05/02/2019
localization_priority: Normal
---


# OLEDBConnection.AlwaysUseConnectionFile property (Excel)

**True** if the connection file is always used to establish a connection to the data source. Read/write **Boolean**.


## Syntax

_expression_.**AlwaysUseConnectionFile**

_expression_ A variable that represents an **[OLEDBConnection](Excel.OLEDBConnection.md)** object.


## Remarks

When this property is **True** the connection file will always be used to establish the connection to the data source. If the connection embedded within the workbook is different from the external connection file, the embedded connection will be ignored and the external connection file will be the only version considered.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]