---
title: Range.RemoveSubtotal method (Excel)
keywords: vbaxl10.chm144185
f1_keywords:
- vbaxl10.chm144185
api_name:
- Excel.Range.RemoveSubtotal
ms.assetid: ec1fd131-551d-009f-1eea-033d805bb34d
ms.date: 05/11/2019
ms.localizationpriority: medium
---


# Range.RemoveSubtotal method (Excel)

Removes subtotals from a list.


## Syntax

_expression_.**RemoveSubtotal**

_expression_ A variable that represents a **[Range](excel.range(object).md)** object.


## Return value

Variant


## Example

This example removes subtotals from the range A1:G37 on Sheet1. The example should be run on a list that has subtotals.

```vb
Worksheets("Sheet1").Range("A1:G37").RemoveSubtotal
```


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]