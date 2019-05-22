---
title: WorksheetFunction.ChiSq_Dist method (Excel)
keywords: vbaxl10.chm137398
f1_keywords:
- vbaxl10.chm137398
ms.prod: excel
api_name:
- Excel.WorksheetFunction.ChiSq_Dist
ms.assetid: be655878-fdb2-7b04-0a9b-6d39652b7e77
ms.date: 05/22/2019
localization_priority: Normal
---


# WorksheetFunction.ChiSq_Dist method (Excel)

Returns the chi-squared distribution.


## Syntax

_expression_.**ChiSq_Dist** (_Arg1_, _Arg2_, _Arg3_)

_expression_ A variable that represents a **[WorksheetFunction](Excel.WorksheetFunction.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Double**|X - The value at which you want to evaluate the distribution.|
| _Arg2_|Required| **Double**|Deg_freedom - The number of degrees of freedom.|
| _Arg3_|Optional| **Variant**|Cumulative - A logical value that determines the form of the function. If cumulative is **True**, **ChiSq_Dist** returns the cumulative distribution function; if **False**, it returns the probability density function. |

## Return value

**Double**


## Remarks

If any argument is nonnumeric, **ChiSq_Dist** returns the #VALUE! error value. 
    
If x is negative, **ChiSq_Dist** returns the #NUM! error value. 
    
If deg_freedom is not an integer, it is truncated. 
    



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]