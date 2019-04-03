---
title: Application.AfterCalculate event (Excel)
keywords: vbaxl10.chm504103
f1_keywords:
- vbaxl10.chm504103
ms.prod: excel
api_name:
- Excel.Application.AfterCalculate
ms.assetid: ed76a36f-1b52-4464-da44-e64c81fb8d38
ms.date: 04/04/2019
localization_priority: Normal
---


# Application.AfterCalculate event (Excel)

The **AfterCalculate** event occurs when all pending refresh activity (both synchronous and asynchronous) and all of the resultant calculation activities have been completed.


## Syntax

_expression_.**AfterCalculate**

_expression_ A variable that represents an **[Application](Excel.Application(object).md)** object.


## Remarks

This event occurs whenever calculation is completed and there are no outstanding queries. It is mandatory for both conditions to be met before the event occurs. The event can be raised even when there is no sheet data in the workbook, such as whenever calculation finishes for the entire workbook and there are no queries running.

Add-in developers use the **AfterCalculate** event to know when all the data in the workbook has been fully updated by any queries and/or calculations that may have been in progress.

This event occurs after all **Worksheet.Calculate**, **Chart.Calculate**, **QueryTable.AfterRefresh**, and **SheetChange** events. It is the last event to occur after all refresh processing and all calc processing have completed, and it occurs after **CalculationState** is set to **xlDone**.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]