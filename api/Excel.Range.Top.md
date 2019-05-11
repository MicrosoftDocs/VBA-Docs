---
title: Range.Top property (Excel)
keywords: vbaxl10.chm144211
f1_keywords:
- vbaxl10.chm144211
ms.prod: excel
api_name:
- Excel.Range.Top
ms.assetid: 0d67ac39-9d35-fc2e-56f1-9bd320a4e8ea
ms.date: 05/11/2019
localization_priority: Normal
---


# Range.Top property (Excel)

Returns a **Variant** value that represents the distance, in [points](../language/glossary/vbe-glossary.md#point), from the top edge of row 1 to the top edge of the range.


## Syntax

_expression_.**Top**

_expression_ A variable that represents a **[Range](excel.range(object).md)** object.


## Remarks

If the range is discontinuous, the first area is used. If the range is more than one row high, the top (lowest numbered) row in the range is used.



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
