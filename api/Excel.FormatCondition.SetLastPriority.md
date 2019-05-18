---
title: FormatCondition.SetLastPriority method (Excel)
keywords: vbaxl10.chm512092
f1_keywords:
- vbaxl10.chm512092
ms.prod: excel
api_name:
- Excel.FormatCondition.SetLastPriority
ms.assetid: fd6263a1-e67f-f4e8-2423-1601f73bdd5c
ms.date: 04/26/2019
localization_priority: Normal
---


# FormatCondition.SetLastPriority method (Excel)

Sets the evaluation order for this conditional formatting rule so that it is evaluated after all other rules on the worksheet.


## Syntax

_expression_.**SetLastPriority**

_expression_ A variable that represents a **[FormatCondition](Excel.FormatCondition.md)** object.


## Remarks

The actual value of the priority will be equal to the total number of conditional formatting rules on the worksheet. When you have multiple conditional formatting rules on a worksheet, this method will cause the priority of rules that had a priority value greater than this rule to be decreased by one.

> [!NOTE] 
> Priority levels for conditional formatting rules are applied on a worksheet-level basis.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]