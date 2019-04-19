---
title: ColorScaleCriterion.Value property (Excel)
keywords: vbaxl10.chm808075
f1_keywords:
- vbaxl10.chm808075
ms.prod: excel
api_name:
- Excel.ColorScaleCriterion.Value
ms.assetid: 829e876f-ca11-855d-bda5-a1c7f86eeb0f
ms.date: 04/20/2019
localization_priority: Normal
---


# ColorScaleCriterion.Value property (Excel)

Returns or sets the minimum, midpoint, or maximum threshold value for a color scale conditional format. Read/write **Variant**.


## Syntax

_expression_.**Value**

_expression_ A variable that represents a **[ColorScaleCriterion](Excel.ColorScaleCriterion.md)** object.


## Remarks

You can set the value only if the **[Type](Excel.ColorScaleCriterion.Type.md)** property for the conditional format is set to one of the following **[XlConditionValueTypes](Excel.XlConditionValueTypes.md)** constants: **xlConditionValueNumber**, **xlConditionValuePercent**, **xlConditionValuePercentile**, or **xlConditionValueFormula**.

If the type of threshold is a formula, you can set the formula as a **String**. The formula must return a single number.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]