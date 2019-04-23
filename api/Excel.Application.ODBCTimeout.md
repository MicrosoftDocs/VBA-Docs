---
title: Application.ODBCTimeout property (Excel)
keywords: vbaxl10.chm133175
f1_keywords:
- vbaxl10.chm133175
ms.prod: excel
api_name:
- Excel.Application.ODBCTimeout
ms.assetid: 92262209-6a0f-f58f-e2d7-2f502f6bd397
ms.date: 04/05/2019
localization_priority: Normal
---


# Application.ODBCTimeout property (Excel)

Returns or sets the ODBC query time limit, in seconds. The default value is 45 seconds. Read/write **Long**.


## Syntax

_expression_.**ODBCTimeout**

_expression_ A variable that represents an **[Application](Excel.Application(object).md)** object.


## Remarks

The value 0 (zero) indicates an indefinite time limit.


## Example

This example sets the ODBC query time limit to 15 seconds.

```vb
Application.ODBCTimeout = 15
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]