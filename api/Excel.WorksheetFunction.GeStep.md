---
title: WorksheetFunction.GeStep method (Excel)
keywords: vbaxl10.chm137296
f1_keywords:
- vbaxl10.chm137296
ms.prod: excel
api_name:
- Excel.WorksheetFunction.GeStep
ms.assetid: dc39a836-c1eb-491f-7f5a-67999c52218a
ms.date: 06/08/2017
localization_priority: Normal
---


# WorksheetFunction.GeStep method (Excel)

Returns 1 if number ? step; returns 0 (zero) otherwise. Use this function to filter a set of values. For example, by summing several GESTEP functions you calculate the count of values that exceed a threshold.


## Syntax

_expression_. `GeStep`( `_Arg1_` , `_Arg2_` )

_expression_ A variable that represents a [WorksheetFunction](./Excel.WorksheetFunction.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Variant**|Number - the value to test against step.|
| _Arg2_|Optional| **Variant**|Step - the threshold value. If you omit a value for step, GESTEP uses zero.|

## Return value

Double


## Remarks

If any argument is nonnumeric, GESTEP returns the #VALUE! error value.


## See also


[WorksheetFunction Object](Excel.WorksheetFunction.md)

