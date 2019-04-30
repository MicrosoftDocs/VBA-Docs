---
title: ODBCConnection.EnableRefresh property (Excel)
keywords: vbaxl10.chm796078
f1_keywords:
- vbaxl10.chm796078
ms.prod: excel
api_name:
- Excel.ODBCConnection.EnableRefresh
ms.assetid: 7d10e758-e92c-90c6-2f12-60b7b5f531ea
ms.date: 05/01/2019
localization_priority: Normal
---


# ODBCConnection.EnableRefresh property (Excel)

**True** if the connection can be refreshed by the user. The default value is **True**. Read/write **Boolean**.


## Syntax

_expression_.**EnableRefresh**

_expression_ A variable that represents an **[ODBCConnection](Excel.ODBCConnection.md)** object.


## Remarks

The **[RefreshOnFileOpen](Excel.ODBCConnection.RefreshOnFileOpen.md)** property is ignored if the **EnableRefresh** property is set to **False**. For OLAP data sources, setting this property to **False** disables updates.



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]