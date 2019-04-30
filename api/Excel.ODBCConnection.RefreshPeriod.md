---
title: ODBCConnection.RefreshPeriod property (Excel)
keywords: vbaxl10.chm796083
f1_keywords:
- vbaxl10.chm796083
ms.prod: excel
api_name:
- Excel.ODBCConnection.RefreshPeriod
ms.assetid: 0e211dad-0ca0-239f-1121-2bae31be2438
ms.date: 05/01/2019
localization_priority: Normal
---


# ODBCConnection.RefreshPeriod property (Excel)

Returns or sets the number of minutes between refreshes. Read/write **Long**.


## Syntax

_expression_.**RefreshPeriod**

_expression_ A variable that represents an **[ODBCConnection](Excel.ODBCConnection.md)** object.


## Remarks

Setting the period to 0 (zero) disables automatic timed refreshes and is equivalent to setting this property to **Null**. The value of the **RefreshPeriod** property can be an integer from 0 through 32767.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]