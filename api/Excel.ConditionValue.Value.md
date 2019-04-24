---
title: ConditionValue.Value property (Excel)
keywords: vbaxl10.chm804075
f1_keywords:
- vbaxl10.chm804075
ms.prod: excel
api_name:
- Excel.ConditionValue.Value
ms.assetid: 376dccc8-2d47-c7ed-1b14-d41dcdd1a8ff
ms.date: 04/23/2019
localization_priority: Normal
---


# ConditionValue.Value property (Excel)

Returns or sets the shortest bar or longest bar threshold value for a data bar conditional format. Read/write **Variant**.


## Syntax

_expression_.**Value**

_expression_ A variable that represents a **[ConditionValue](Excel.ConditionValue.md)** object.


## Remarks

You can set the value only if the **[Type](Excel.ConditionValue.Type.md)** property for the conditional format is set to one of the following **[XlConditionValueTypes](Excel.XlConditionValueTypes.md)** constants: **xlConditionValueNumber**, **xlConditionValuePercent**, **xlConditionValuePercentile**, or **xlConditionValueFormula**.

If the threshold type is a formula, you can set the formula as a **String**. The formula must return a single number.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]