---
title: UniqueValues.SetFirstPriority method (Excel)
keywords: vbaxl10.chm826082
f1_keywords:
- vbaxl10.chm826082
ms.prod: excel
api_name:
- Excel.UniqueValues.SetFirstPriority
ms.assetid: 65e0be2a-1bc2-167d-516f-3ba0ebab1322
ms.date: 05/18/2019
localization_priority: Normal
---


# UniqueValues.SetFirstPriority method (Excel)

Sets the priority value for this conditional formatting rule to 1 so that it will be evaluated before all other rules on the worksheet.


## Syntax

_expression_.**SetFirstPriority**

_expression_ A variable that represents a **[UniqueValues](Excel.UniqueValues.md)** object.


## Remarks

When you have multiple conditional formatting rules on a worksheet, if the rule was not previously set to priority 1, this method causes the priority of all other existing rules on the worksheet to be increased by one.

> [!NOTE] 
> Priority levels for conditional formatting rules are applied on a worksheet-level basis.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]