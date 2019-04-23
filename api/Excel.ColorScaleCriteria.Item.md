---
title: ColorScaleCriteria.Item property (Excel)
keywords: vbaxl10.chm807076
f1_keywords:
- vbaxl10.chm807076
ms.prod: excel
api_name:
- Excel.ColorScaleCriteria.Item
ms.assetid: 62033ea0-19c6-430f-0b9e-9eae62791352
ms.date: 04/20/2019
localization_priority: Normal
---


# ColorScaleCriteria.Item property (Excel)

Returns a single **[ColorScaleCriterion](Excel.ColorScaleCriterion.md)** object from the **ColorScaleCriteria** collection. Read-only.


## Syntax

_expression_.**Item** (_Index_)

_expression_ A variable that represents a **[ColorScaleCriteria](Excel.ColorScaleCriteria.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Variant**|The index number of the **ColorScaleCriterion** object.|

## Remarks

The value of the _Index_ parameter cannot be greater than the number of criteria set for a color scale conditional format. The criteria are equivalent to the threshold values assigned for the color scale. 

To find the number of threshold values, use the **[Count](Excel.ColorScaleCriteria.Count.md)** property.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]