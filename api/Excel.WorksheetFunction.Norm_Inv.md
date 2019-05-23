---
title: WorksheetFunction.Norm_Inv method (Excel)
keywords: vbaxl10.chm137371
f1_keywords:
- vbaxl10.chm137371
ms.prod: excel
api_name:
- Excel.WorksheetFunction.Norm_Inv
ms.assetid: 0069b45f-629d-6212-18da-6954be00181f
ms.date: 05/24/2019
localization_priority: Normal
---


# WorksheetFunction.Norm_Inv method (Excel)

Returns the inverse of the normal cumulative distribution for the specified mean and standard deviation.


## Syntax

_expression_.**Norm_Inv** (_Arg1_, _Arg2_, _Arg3_)

_expression_ A variable that represents a **[WorksheetFunction](Excel.WorksheetFunction.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Double**|Probability - A probability corresponding to the normal distribution.|
| _Arg2_|Required| **Double**|Mean - The arithmetic mean of the distribution.|
| _Arg3_|Required| **Double**|Standard_dev - The standard deviation of the distribution.|

## Return value

**Double**


## Remarks

If any argument is non-numeric, **Norm_Inv** returns the #VALUE! error value.
    
If probability ≤ 0 or if probability ≥ 1, **Norm_Inv** returns the #NUM! error value.
    
If standard_dev ≤ 0, **Norm_Inv** returns the #NUM! error value.
    
If mean = 0 and standard_dev = 1, **Norm_Inv** uses the standard normal distribution (see **[Norm_S_Inv](excel.worksheetfunction.norm_s_inv.md)**).
    
Given a value for probability, **Norm_Inv** seeks that value x such that NORM_DIST(x, mean, standard_dev, TRUE) = probability. Thus, precision of **Norm_Inv** depends on precision of **Norm_Dist**.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]