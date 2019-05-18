---
title: AboveAverage.SetFirstPriority method (Excel)
keywords: vbaxl10.chm824082
f1_keywords:
- vbaxl10.chm824082
ms.prod: excel
api_name:
- Excel.AboveAverage.SetFirstPriority
ms.assetid: 4f9b02ff-232b-3dcb-239b-6ba7897366d0
ms.date: 03/26/2019
localization_priority: Normal
---


# AboveAverage.SetFirstPriority method (Excel)

Sets the priority value for this conditional formatting rule to "1" so that it will be evaluated before all other rules on the worksheet.


## Syntax

_expression_.**SetFirstPriority**

_expression_ A variable that represents an **[AboveAverage](Excel.AboveAverage.md)** object.


## Remarks

When you have multiple conditional formatting rules on a worksheet, if the rule was not previously set to priority "1", this method will cause the priority of all other existing rules on the worksheet to be increased by one.

> [!NOTE] 
> Priority levels for conditional formatting rules are applied on a worksheet-level basis.



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]