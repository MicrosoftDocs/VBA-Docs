---
title: Range.QueryTable property (Excel)
keywords: vbaxl10.chm144183
f1_keywords:
- vbaxl10.chm144183
ms.prod: excel
api_name:
- Excel.Range.QueryTable
ms.assetid: 6370d43c-74b5-1bb9-f849-c70006432504
ms.date: 05/11/2019
localization_priority: Normal
---


# Range.QueryTable property (Excel)

Returns a **[QueryTable](Excel.QueryTable.md)** object that represents the query table that intersects the specified **Range** object.


## Syntax

_expression_.**QueryTable**

_expression_ A variable that represents a **[Range](excel.range(object).md)** object.


## Example

This example refreshes the **QueryTable** object that intersects cell A10 on worksheet one.

```vb
Worksheets(1).Range("a10").QueryTable.Refresh
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]