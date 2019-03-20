---
title: FormatCondition.ShortestBarValue property (Access)
keywords: vbaac10.chm14531
f1_keywords:
- vbaac10.chm14531
ms.prod: access
api_name:
- Access.FormatCondition.ShortestBarValue
ms.assetid: b262c385-0c12-87cc-45cc-83a658a61510
ms.date: 03/20/2019
localization_priority: Normal
---


# FormatCondition.ShortestBarValue property (Access)

Gets or sets a numeric expression that specifies the value of the shortest bar of a **FormatCondition**. Read/write **String**.


## Syntax

_expression_.**ShortestBarValue**

_expression_ A variable that represents a **[FormatCondition](Access.FormatCondition.md)** object.


## Remarks

By default, the **ShortestBarValue** contains a zero-length string ("").

If the value of the **[ShortestBarLimit](Access.FormatCondition.ShortestBarLimit.md)** property is **acAutomatic**, the **ShortestBarValue** is ignored.



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]