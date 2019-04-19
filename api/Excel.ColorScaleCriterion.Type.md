---
title: ColorScaleCriterion.Type property (Excel)
keywords: vbaxl10.chm808074
f1_keywords:
- vbaxl10.chm808074
ms.prod: excel
api_name:
- Excel.ColorScaleCriterion.Type
ms.assetid: 59ea77b7-4d12-22e5-380c-bb94912a6550
ms.date: 04/20/2019
localization_priority: Normal
---


# ColorScaleCriterion.Type property (Excel)

Returns one of the constants of the **[XlConditionValueTypes](Excel.XlConditionValueTypes.md)** enumeration, which specifies how the threshold values for a data bar or color scale conditional format are determined. Read-only.


## Syntax

_expression_.**Type**

_expression_ A variable that represents a **[ColorScaleCriterion](Excel.ColorScaleCriterion.md)** object.


## Remarks

The type of threshold value for a data bar or color scale can be a number, percent, formula, or percentile. Setting the type to percentile will use the Percentile function in Microsoft Excel to determine the threshold value.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]