---
title: UniqueValues.Priority property (Excel)
keywords: vbaxl10.chm826073
f1_keywords:
- vbaxl10.chm826073
api_name:
- Excel.UniqueValues.Priority
ms.assetid: c966b2b0-76ed-4ff4-0abc-950ffc7209e1
ms.date: 05/18/2019
ms.localizationpriority: medium
---


# UniqueValues.Priority property (Excel)

Returns or sets the priority value of the conditional formatting rule. The priority determines the order of evaluation when multiple conditional formatting rules exist on a worksheet.


## Syntax

_expression_.**Priority**

_expression_ A variable that represents a **[UniqueValues](Excel.UniqueValues.md)** object.


## Remarks

When setting the priority, the value must be a positive integer between 1 and the total number of conditional formatting rules on the worksheet. The priority must be a unique value for all rules on the worksheet, so changing the priority for the specified conditional formatting rule may cause the priority value of the other rules on the worksheet to be shifted.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]