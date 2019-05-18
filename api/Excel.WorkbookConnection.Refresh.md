---
title: WorkbookConnection.Refresh method (Excel)
keywords: vbaxl10.chm774081
f1_keywords:
- vbaxl10.chm774081
ms.prod: excel
api_name:
- Excel.WorkbookConnection.Refresh
ms.assetid: 5e6f045f-6625-857c-eb55-ac52f70e8fb9
ms.date: 05/18/2019
localization_priority: Normal
---


# WorkbookConnection.Refresh method (Excel)

Refreshes a workbook connection.


## Syntax

_expression_.**Refresh**

_expression_ A variable that represents a **[WorkbookConnection](Excel.WorkbookConnection.md)** object.


## Remarks

If the **[DisplayAlerts](Excel.Application.DisplayAlerts.md)** property is **False**, dialog boxes are not displayed, and the **Refresh** method fails with the Insufficient Connection Information exception.

A refresh failure for one connection will not have any impact on refresh operations for the other connections.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]