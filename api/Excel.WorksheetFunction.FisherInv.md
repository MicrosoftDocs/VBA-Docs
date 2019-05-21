---
title: WorksheetFunction.FisherInv method (Excel)
keywords: vbaxl10.chm137188
f1_keywords:
- vbaxl10.chm137188
ms.prod: excel
api_name:
- Excel.WorksheetFunction.FisherInv
ms.assetid: bf4656e3-b79d-7fe6-917f-16afedc736fe
ms.date: 05/22/2019
localization_priority: Normal
---


# WorksheetFunction.FisherInv method (Excel)

Returns the inverse of the Fisher transformation. Use this transformation when analyzing correlations between ranges or arrays of data. If y = FISHER(x), then FISHERINV(y) = x.


## Syntax

_expression_.**FisherInv** (_Arg1_)

_expression_ A variable that represents a **[WorksheetFunction](Excel.WorksheetFunction.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Double**|y - the value for which you want to perform the inverse of the transformation.|

## Return value

**Double**


## Remarks

If y is nonnumeric, **FisherInv** returns the #VALUE! error value.
    
The equation for the inverse of the **Fisher** transformation is ![Formula](../images/awffshri_ZA06051142.gif)




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]