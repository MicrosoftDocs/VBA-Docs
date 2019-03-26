---
title: AboveAverage.CalcFor property (Excel)
keywords: vbaxl10.chm824088
f1_keywords:
- vbaxl10.chm824088
ms.prod: excel
api_name:
- Excel.AboveAverage.CalcFor
ms.assetid: 9a9e04df-f3f8-2daa-b58c-3245f4bfe6c9
ms.date: 03/26/2019
localization_priority: Normal
---


# AboveAverage.CalcFor property (Excel)

Returns or sets one of the constants of the **[XlCalcFor](Excel.XlCalcFor.md)** enumeration, which specifies the scope of data to be evaluated for the conditional format in a PivotTable report.


## Syntax

_expression_.**CalcFor**

_expression_ A variable that represents an **[AboveAverage](Excel.AboveAverage.md)** object.


## Remarks

This property is applicable only when the conditional format is being applied to data in a PivotTable report.

This property can be set to **xlAllValues**, **xlColGroups**, or **xlRowGroups** only if the **[AboveAverage.ScopeType](Excel.AboveAverage.ScopeType.md)** property is set to **xlFieldsScope**.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]