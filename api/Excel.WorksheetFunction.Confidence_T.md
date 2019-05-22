---
title: WorksheetFunction.Confidence_T method (Excel)
keywords: vbaxl10.chm137360
f1_keywords:
- vbaxl10.chm137360
ms.prod: excel
api_name:
- Excel.WorksheetFunction.Confidence_T
ms.assetid: b4e497b6-bf5a-5630-3092-d806012e0c97
ms.date: 05/22/2019
localization_priority: Normal
---


# WorksheetFunction.Confidence_T method (Excel)

Returns the confidence interval for a population mean, using a Student's t distribution.


## Syntax

_expression_.**Confidence_T** (_Arg1_, _Arg2_, _Arg3_)

_expression_ A variable that represents a **[WorksheetFunction](Excel.WorksheetFunction.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Double**|Alpha - The significance level used to compute the confidence level. The confidence level equals 100*(1 - alpha)%, or in other words, an alpha of 0.05 indicates a 95 percent confidence level.|
| _Arg2_|Required| **Double**|Standard_dev - The population standard deviation for the data range; is assumed to be known.|
| _Arg3_|Required| **Double**|Size - The sample size.|

## Return value

**Double**


## Remarks

If any argument is nonnumeric, **Confidence_T** returns the #VALUE! error value. 
    
If alpha ≤ 0 or alpha ≥ 1, **Confidence_T** returns the #NUM! error value. 
    
If standard_dev ≤ 0, **Confidence_T** returns the #NUM! error value. 
    
If size is not an integer, it is truncated. 
    
If size equals 1, **Confidence_T** returns the #DIV/0! error value.  




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]