---
title: OLEDBConnection.RefreshPeriod property (Excel)
keywords: vbaxl10.chm794087
f1_keywords:
- vbaxl10.chm794087
ms.prod: excel
api_name:
- Excel.OLEDBConnection.RefreshPeriod
ms.assetid: 23032291-8491-42b8-b0ee-18287c115c29
ms.date: 05/02/2019
localization_priority: Normal
---


# OLEDBConnection.RefreshPeriod property (Excel)

Returns or sets the number of minutes between refreshes. Read/write **Long**.


## Syntax

_expression_.**RefreshPeriod**

_expression_ A variable that represents an **[OLEDBConnection](Excel.OLEDBConnection.md)** object.


## Remarks

Setting the period to 0 (zero) disables automatic timed refreshes and is equivalent to setting this property to **Null**. The value of the **RefreshPeriod** property can be an integer from 0 through 32767.



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]