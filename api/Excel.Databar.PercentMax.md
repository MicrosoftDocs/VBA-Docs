---
title: DataBar.PercentMax property (Excel)
keywords: vbaxl10.chm810079
f1_keywords:
- vbaxl10.chm810079
api_name:
- Excel.DataBar.PercentMax
ms.assetid: d06a5ce2-a298-7974-f9bc-f8fb3fd7ccf0
ms.date: 04/23/2019
ms.localizationpriority: medium
---


# DataBar.PercentMax property (Excel)

Returns or sets a **Long** value that specifies the length of the longest data bar as a percentage of cell width.


## Syntax

_expression_.**PercentMax**

_expression_ A variable that represents a **[DataBar](Excel.DataBar.md)** object.


## Remarks

The value must be a whole number between 0 and 100. The default value is 100.

The effect of the **PercentMax** property varies depending on the setting of the **[AxisPosition](Excel.DataBar.AxisPosition.md)** property of the **DataBar** object. 

When the **AxisPosition** property is **xlDataBarAxisAutomatic** and the range contains both positive and negative values, the sum of the lengths of the longest positive data bar and the longest negative data bar will not exceed the value specified by the **PercentMax** property. 

When the **AxisPosition** property is **xlDataBarAxisMidpoint**, the longest data bar (positive or negative) will be equal to the value of the **PercentMax** property divided by 2. 

When the **AxisPosition** property is **xlDataBarAxisNone**, the length of the longest data bar is always the percentage of the cell width specified by the **PercentMax** property.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]