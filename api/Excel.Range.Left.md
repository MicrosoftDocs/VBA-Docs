---
title: Range.Left property (Excel)
keywords: vbaxl10.chm144153
f1_keywords:
- vbaxl10.chm144153
ms.prod: excel
api_name:
- Excel.Range.Left
ms.assetid: 634fa7b8-3565-6178-f564-3c5d24c16d26
ms.date: 06/08/2017
localization_priority: Priority
---


# Range.Left property (Excel)

Returns a  **Variant** value that represents the distance, in points, from the left edge of column A to the left edge of the range.


## Syntax

_expression_. `Left`

_expression_ A variable that represents a [Range](excel.range-graph-property.md) object.


## Remarks

If the range is discontinuous, the first area is used. If the range is more than one column wide, the leftmost column in the range is used.


## See also


[Range Object](Excel.Range(object).md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]