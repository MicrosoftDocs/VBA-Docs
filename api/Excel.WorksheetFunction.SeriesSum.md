---
title: WorksheetFunction.SeriesSum method (Excel)
keywords: vbaxl10.chm137291
f1_keywords:
- vbaxl10.chm137291
ms.prod: excel
api_name:
- Excel.WorksheetFunction.SeriesSum
ms.assetid: 096faaa8-4bd3-fd61-4442-b29785a93c7c
ms.date: 05/25/2019
localization_priority: Normal
---


# WorksheetFunction.SeriesSum method (Excel)

Returns the sum of a power series based on the following formula:

> ![Formula](../images/awfsrssm_ZA06051246.gif)


## Syntax

_expression_.**SeriesSum** (_Arg1_, _Arg2_, _Arg3_, _Arg4_)

_expression_ A variable that represents a **[WorksheetFunction](Excel.WorksheetFunction.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Variant**|X - the input value to the power series.|
| _Arg2_|Required| **Variant**|N - the initial power to which you want to raise x.|
| _Arg3_|Required| **Variant**|M - the step by which to increase n for each term in the series.|
| _Arg4_|Required| **Variant**|Coefficients - a set of coefficients by which each successive power of x is multiplied. The number of values in coefficients determines the number of terms in the power series. For example, if there are three values in coefficients, there will be three terms in the power series.|

## Return value

**Double**


## Remarks

If any argument is nonnumeric, **SeriesSum** returns the #VALUE! error value.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]