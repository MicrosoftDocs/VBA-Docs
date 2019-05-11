---
title: Range.ShrinkToFit property (Excel)
keywords: vbaxl10.chm144199
f1_keywords:
- vbaxl10.chm144199
ms.prod: excel
api_name:
- Excel.Range.ShrinkToFit
ms.assetid: fc9aed64-1000-3419-ceb7-a95c15f8a2d0
ms.date: 05/11/2019
localization_priority: Normal
---


# Range.ShrinkToFit property (Excel)

Returns or sets a **Variant** value that indicates if text automatically shrinks to fit in the available column width.


## Syntax

_expression_.**ShrinkToFit**

_expression_ A variable that represents a **[Range](excel.range(object).md)** object.


## Remarks

This property returns **True** if text automatically shrinks to fit in the available column width, or **Null** if this property isn't set to the same value for all cells in the specified range.


## Example

This example causes text in row one to automatically shrink to fit in the available column width.

```vb
Rows(1).ShrinkToFit = True
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]