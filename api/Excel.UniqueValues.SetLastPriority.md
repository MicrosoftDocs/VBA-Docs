---
title: UniqueValues.SetLastPriority method (Excel)
keywords: vbaxl10.chm826083
f1_keywords:
- vbaxl10.chm826083
ms.prod: excel
api_name:
- Excel.UniqueValues.SetLastPriority
ms.assetid: 9e2db204-4a9f-1690-7fc1-bec371fccaff
ms.date: 05/18/2019
localization_priority: Normal
---


# UniqueValues.SetLastPriority method (Excel)

Sets the evaluation order for this conditional formatting rule so that it is evaluated after all other rules on the worksheet.


## Syntax

_expression_.**SetLastPriority**

_expression_ A variable that represents a **[UniqueValues](Excel.UniqueValues.md)** object.


## Remarks

The actual value of the priority will be equal to the total number of conditional formatting rules on the worksheet. When you have multiple conditional formatting rules on a worksheet, this method causes the priority of rules that had a priority value greater than this rule to be decreased by one.

> [!NOTE] 
> Priority levels for conditional formatting rules are applied on a worksheet-level basis.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]