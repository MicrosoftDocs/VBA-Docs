---
title: ColorScale.Priority property (Excel)
keywords: vbaxl10.chm806073
f1_keywords:
- vbaxl10.chm806073
ms.prod: excel
api_name:
- Excel.ColorScale.Priority
ms.assetid: 71209666-9c7c-f376-69cd-bc73c38a9681
ms.date: 04/20/2019
localization_priority: Normal
---


# ColorScale.Priority property (Excel)

Returns or sets the priority value of the conditional formatting rule. The priority determines the order of evaluation when multiple conditional formatting rules exist on a worksheet.


## Syntax

_expression_.**Priority**

_expression_ A variable that represents a **[ColorScale](Excel.ColorScale.md)** object.


## Remarks

When setting the priority, the value must be a positive integer between 1 and the total number of conditional formatting rules on the worksheet. The priority must be a unique value for all rules on the worksheet, so changing the priority for the specified conditional formatting rule may cause the priority value of the other rules on the worksheet to be shifted.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]