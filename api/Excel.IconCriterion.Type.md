---
title: IconCriterion.Type property (Excel)
keywords: vbaxl10.chm814074
f1_keywords:
- vbaxl10.chm814074
ms.prod: excel
api_name:
- Excel.IconCriterion.Type
ms.assetid: bbe75bbb-42d1-7b71-7a7a-7c51e8c47cbc
ms.date: 04/27/2019
localization_priority: Normal
---


# IconCriterion.Type property (Excel)

Returns one of the constants of the **[XlConditionValueTypes](Excel.XlConditionValueTypes.md)** enumeration, which specifies how the threshold value for an icon set is determined. Read-only.


## Syntax

_expression_.**Type**

_expression_ A variable that represents an **[IconCriterion](Excel.IconCriterion.md)** object.


## Remarks

The type of threshold value for an icon set can be a number, percent, formula, or percentile. Setting the type to percentile will use the **Percentile** function in Excel to determine the threshold value.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]