---
title: RTD.RefreshData method (Excel)
keywords: vbaxl10.chm728074
f1_keywords:
- vbaxl10.chm728074
ms.prod: excel
api_name:
- Excel.RTD.RefreshData
ms.assetid: fa2ddf47-1821-25b6-fcd9-b42853c2689a
ms.date: 05/11/2019
localization_priority: Normal
---


# RTD.RefreshData method (Excel)

Requests an update of real-time data from the real-time data server.


## Syntax

_expression_.**RefreshData**

_expression_ A variable that represents an **[RTD](Excel.RTD.md)** object.


## Remarks

Avoid using the **RefreshData** method in user-defined functions because this method will fail if it is called during recalculation.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]