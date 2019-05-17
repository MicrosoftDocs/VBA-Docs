---
title: Top10.CalcFor property (Excel)
keywords: vbaxl10.chm822090
f1_keywords:
- vbaxl10.chm822090
ms.prod: excel
api_name:
- Excel.Top10.CalcFor
ms.assetid: 4ab81649-8221-a83d-5e51-0c4dbe01b175
ms.date: 05/18/2019
localization_priority: Normal
---


# Top10.CalcFor property (Excel)

Returns or sets one of the constants of the **[XlCalcFor](Excel.XlCalcFor.md)** enumeration, which specifies how the conditional format in a PivotTable report should be evaluated.


## Syntax

_expression_.**CalcFor**

_expression_ A variable that represents a **[Top10](Excel.Top10.md)** object.


## Remarks

This property is applicable only when the conditional format is being applied to data in a PivotTable report.

This property can be set to **xlAllValues**, **xlColGroups**, or **xlRowGroups** only if the **[ScopeType](Excel.Top10.ScopeType.md)** property is set to **xlFieldsScope**.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]