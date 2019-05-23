---
title: WorksheetFunction.FactDouble method (Excel)
keywords: vbaxl10.chm137292
f1_keywords:
- vbaxl10.chm137292
ms.prod: excel
api_name:
- Excel.WorksheetFunction.FactDouble
ms.assetid: 71d8d537-b06c-7614-d6d6-b6c57ed8c68f
ms.date: 05/22/2019
localization_priority: Normal
---


# WorksheetFunction.FactDouble method (Excel)

Returns the double factorial of a number.


## Syntax

_expression_.**FactDouble** (_Arg1_)

_expression_ A variable that represents a **[WorksheetFunction](Excel.WorksheetFunction.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Variant**|Number - the value for which to return the double factorial. If number is not an integer, it is truncated.|

## Return value

**Double**


## Remarks

If number is nonnumeric, **FactDouble** returns the #VALUE! error value.
    
If number is negative, **FactDouble** returns the #NUM! error value.
    
If number is even: &nbsp; ![Formula](../images/awffdbl1_ZA06051139.gif)

If number is odd: &nbsp; ![Formula](../images/awffdbl2_ZA06051140.gif)




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]