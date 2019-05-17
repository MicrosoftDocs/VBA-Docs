---
title: ColorScale.SetFirstPriority method (Excel)
keywords: vbaxl10.chm806078
f1_keywords:
- vbaxl10.chm806078
ms.prod: excel
api_name:
- Excel.ColorScale.SetFirstPriority
ms.assetid: 812bf48e-066c-6bea-be43-1a068c948ea8
ms.date: 04/20/2019
localization_priority: Normal
---


# ColorScale.SetFirstPriority method (Excel)

Sets the priority value for this conditional formatting rule to "1" so that it will be evaluated before all other rules on the worksheet.


## Syntax

_expression_.**SetFirstPriority**

_expression_ A variable that represents a **[ColorScale](Excel.ColorScale.md)** object.


## Remarks

When you have multiple conditional formatting rules on a worksheet, if the rule was not previously set to priority "1", this method will cause the priority of all other existing rules on the worksheet to be increased by one.

> [!NOTE] 
> Priority levels for conditional formatting rules are applied on a worksheet-level basis.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]