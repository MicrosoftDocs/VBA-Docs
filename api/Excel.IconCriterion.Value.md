---
title: IconCriterion.Value property (Excel)
keywords: vbaxl10.chm814075
f1_keywords:
- vbaxl10.chm814075
ms.prod: excel
api_name:
- Excel.IconCriterion.Value
ms.assetid: 5cb72b0b-1df2-dd47-932f-1454fda9f804
ms.date: 04/27/2019
localization_priority: Normal
---


# IconCriterion.Value property (Excel)

Returns or sets the threshold value for an icon in a conditional format. Read/write **Variant**.


## Syntax

_expression_.**Value**

_expression_ A variable that represents an **[IconCriterion](Excel.IconCriterion.md)** object.


## Remarks

You can set the value only if the **[Type](Excel.IconCriterion.Type.md)** property for the conditional format is set to one of the following **[XlConditionValueTypes](Excel.XlConditionValueTypes.md)** constants: **xlConditionValueNumber**, **xlConditionValuePercent**, **xlConditionValuePercentile**, or **xlConditionValueFormula**.

If the type of threshold is a formula, you can set the formula as a **String**. The formula must return a single number.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]