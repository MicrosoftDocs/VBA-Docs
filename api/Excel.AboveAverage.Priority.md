---
title: AboveAverage.Priority property (Excel)
keywords: vbaxl10.chm824073
f1_keywords:
- vbaxl10.chm824073
api_name:
- Excel.AboveAverage.Priority
ms.assetid: 4df00b9f-d260-8b1b-de08-0886bdc87a1c
ms.date: 03/26/2019
ms.localizationpriority: medium
---


# AboveAverage.Priority property (Excel)

Returns or sets the priority value of the conditional formatting rule. The priority determines the order of evaluation when multiple conditional formatting rules exist on a worksheet.


## Syntax

_expression_.**Priority**

_expression_ A variable that represents an **[AboveAverage](Excel.AboveAverage.md)** object.


## Remarks

When setting the priority, the value must be a positive integer between 1 and the total number of conditional formatting rules on the worksheet. The priority must be a unique value for all rules on the worksheet, so changing the priority for the specified conditional formatting rule may cause the priority value of the other rules on the worksheet to be shifted.



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]