---
title: ColorScale.SetLastPriority method (Excel)
keywords: vbaxl10.chm806079
f1_keywords:
- vbaxl10.chm806079
ms.prod: excel
api_name:
- Excel.ColorScale.SetLastPriority
ms.assetid: 01c64e4d-98e8-3647-5e06-23fd1000757b
ms.date: 04/20/2019
localization_priority: Normal
---


# ColorScale.SetLastPriority method (Excel)

Sets the evaluation order for this conditional formatting rule so that it is evaluated after all other rules on the worksheet.


## Syntax

_expression_.**SetLastPriority**

_expression_ A variable that represents a **[ColorScale](Excel.ColorScale.md)** object.


## Remarks

The actual value of the priority will be equal to the total number of conditional formatting rules on the worksheet. When you have multiple conditional formatting rules on a worksheet, this method will cause the priority of rules that had a priority value greater than this rule to be decreased by one.

> [!NOTE] 
> Priority levels for conditional formatting rules are applied on a worksheet-level basis.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]