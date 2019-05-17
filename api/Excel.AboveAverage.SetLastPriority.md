---
title: AboveAverage.SetLastPriority method (Excel)
keywords: vbaxl10.chm824083
f1_keywords:
- vbaxl10.chm824083
ms.prod: excel
api_name:
- Excel.AboveAverage.SetLastPriority
ms.assetid: e28605d2-338b-4efb-e7f0-f250bca85050
ms.date: 03/26/2019
localization_priority: Normal
---


# AboveAverage.SetLastPriority method (Excel)

Sets the evaluation order for this conditional formatting rule so that it is evaluated after all other rules on the worksheet.


## Syntax

_expression_.**SetLastPriority**

_expression_ A variable that represents an **[AboveAverage](Excel.AboveAverage.md)** object.


## Remarks

The actual value of the priority will be equal to the total number of conditional formatting rules on the worksheet. When you have multiple conditional formatting rules on a worksheet, this method will cause the priority of rules that had a priority value greater than this rule to be decreased by one.

> [!NOTE] 
> Priority levels for conditional formatting rules are applied on a worksheet-level basis.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]