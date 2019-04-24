---
title: DataBar.PercentMin property (Excel)
keywords: vbaxl10.chm810078
f1_keywords:
- vbaxl10.chm810078
ms.prod: excel
api_name:
- Excel.DataBar.PercentMin
ms.assetid: bd8670f9-ae0b-3a1c-5b14-84cc00638b6e
ms.date: 04/23/2019
localization_priority: Normal
---


# DataBar.PercentMin property (Excel)

Returns or sets a **Long** value that specifies the length of the shortest data bar as a percentage of cell width.


## Syntax

_expression_.**PercentMin**

_expression_ A variable that represents a **[DataBar](Excel.DataBar.md)** object.


## Remarks

The value must be a whole number between 0 and 100. The default value is 0.

The effect of the **PercentMin** property varies depending on the setting of the **[AxisPosition](Excel.DataBar.AxisPosition.md)** property of the **DataBar** object. 

When the **AxisPosition** property is **xlDataBarAxisAutomatic** and the range contains both positive and negative values, the minimum length of a positive or negative bar is specified by the **PercentMin** property, and the axis is displayed by using automatic centering rules. 

When the **AxisPosition** property is **xlDataBarAxisMidpoint**, the minimum length of a positive or negative bar is specified by the **PercentMin** property, and the axis is centered in the middle of the cell. 

When the **AxisPosition** property is **xlDataBarAxisNone**, the length of the shortest data bar is always the percentage of the cell width specified by the **PercentMin** property.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]