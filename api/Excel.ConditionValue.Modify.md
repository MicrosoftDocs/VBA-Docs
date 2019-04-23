---
title: ConditionValue.Modify method (Excel)
keywords: vbaxl10.chm804073
f1_keywords:
- vbaxl10.chm804073
ms.prod: excel
api_name:
- Excel.ConditionValue.Modify
ms.assetid: 3da6d850-7b7b-2419-b211-b18081c31e77
ms.date: 04/23/2019
localization_priority: Normal
---


# ConditionValue.Modify method (Excel)

Modifies how the longest bar or shortest bar is evaluated for a data bar conditional formatting rule. 


## Syntax

_expression_.**Modify** (_NewType_, _NewValue_)

_expression_ A variable that represents a **[ConditionValue](Excel.ConditionValue.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _NewType_|Required| **[XlConditionValueTypes](Excel.XlConditionValueTypes.md)**|Specifies how the shortest bar or longest bar is evaluated. The default value is **xlConditionLowestValue** for the shortest bar and **xlConditionHighestValue** for the longest bar.|
| _NewValue_|Optional| **Variant**|The value assigned to the shortest or longest data bar. Depending on the _NewType_ argument, this can be a number or a formula that evaluates to a number.|

## Remarks

The following table describes the acceptable threshold values for each type of evaluation.

|_NewType_ argument|_NewValue_ argument|
|:-----|:-----|
|xlConditionLowestValue|Argument is ignored.|
|xlConditionHighestValue|Argument is ignored.|
|xlConditionValueNumber|Any number.|
|xlConditionValuePercent|Any number between 0 and 100. |
|xlConditionValuePercentile|Any number between 0 and 100.|
|xlConditionValueFormula|A formula that returns a single number.|



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]