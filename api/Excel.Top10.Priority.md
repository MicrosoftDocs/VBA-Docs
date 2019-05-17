---
title: Top10.Priority property (Excel)
keywords: vbaxl10.chm822073
f1_keywords:
- vbaxl10.chm822073
ms.prod: excel
api_name:
- Excel.Top10.Priority
ms.assetid: 0f54585a-2390-dfde-d4c2-5f0c1e9f8ff7
ms.date: 05/18/2019
localization_priority: Normal
---


# Top10.Priority property (Excel)

Returns or sets the priority value of the conditional formatting rule. The priority determines the order of evaluation when multiple conditional formatting rules exist on a worksheet.


## Syntax

_expression_.**Priority**

_expression_ A variable that represents a **[Top10](Excel.Top10.md)** object.


## Remarks

When setting the priority, the value must be a positive integer between 1 and the total number of conditional formatting rules on the worksheet. The priority must be a unique value for all rules on the worksheet, so changing the priority for the specified conditional formatting rule may cause the priority value of the other rules on the worksheet to be shifted.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]