---
title: FormatCondition.LongestBarValue property (Access)
keywords: vbaac10.chm14533
f1_keywords:
- vbaac10.chm14533
ms.prod: access
api_name:
- Access.FormatCondition.LongestBarValue
ms.assetid: bff378d6-138a-31bf-4587-d0f4ce827240
ms.date: 03/20/2019
localization_priority: Normal
---


# FormatCondition.LongestBarValue property (Access)

Gets or sets a numeric expression that specifies the value of the longest bar of a **FormatCondition**. Read/write **String**.


## Syntax

_expression_.**LongestBarValue**

_expression_ A variable that represents a **[FormatCondition](Access.FormatCondition.md)** object.


## Remarks

By default, the **LongestBarValue** contains a zero-length string ("").

If the value of the **[LongestBarLimit](Access.FormatCondition.LongestBarLimit.md)** property is **acAutomatic**, the **LongestBarValue** is ignored.



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]