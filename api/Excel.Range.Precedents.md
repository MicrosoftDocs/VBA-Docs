---
title: Range.Precedents property (Excel)
keywords: vbaxl10.chm144178
f1_keywords:
- vbaxl10.chm144178
ms.prod: excel
api_name:
- Excel.Range.Precedents
ms.assetid: 3c00cfb4-1c12-668d-a952-89f9b1ef129f
ms.date: 05/11/2019
localization_priority: Normal
---


# Range.Precedents property (Excel)

Returns a **Range** object that represents all the precedents of a cell. This can be a multiple selection (a union of **Range** objects) if there's more than one precedent. Read-only.


## Syntax

_expression_.**Precedents**

_expression_ A variable that represents a **[Range](excel.range(object).md)** object.


## Example

This example selects the precedents of cell A1 on Sheet1.

```vb
Worksheets("Sheet1").Activate 
Range("A1").Precedents.Select
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]