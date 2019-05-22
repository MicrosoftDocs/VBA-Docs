---
title: WorksheetFunction.F_Test method (Excel)
keywords: vbaxl10.chm137362
f1_keywords:
- vbaxl10.chm137362
ms.prod: excel
api_name:
- Excel.WorksheetFunction.F_Test
ms.assetid: 193fefdf-28f9-6635-19ec-10c8f655eaf1
ms.date: 05/22/2019
localization_priority: Normal
---


# WorksheetFunction.F_Test method (Excel)

Returns the result of an F-test. An F-test returns the two-tailed probability that the variances in array1 and array2 are not significantly different. Use this function to determine whether two samples have different variances. For example, given test scores from public and private schools, you can test whether these schools have different levels of test score diversity.


## Syntax

_expression_.**F_Test** (_Arg1_, _Arg2_)

_expression_ A variable that represents a **[WorksheetFunction](Excel.WorksheetFunction.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Variant**|Array1 - the first array or range of data.|
| _Arg2_|Required| **Variant**|Array2 - the second array or range of data.|

## Return value

**Double**


## Remarks

The arguments must be either numbers or names, arrays, or references that contain numbers.
    
If an array or reference argument contains text, logical values, or empty cells, those values are ignored; however, cells with the value zero are included.
    
If the number of data points in array1 or array2 is less than 2, or if the variance of array1 or array2 is zero, **F_Test** returns the #DIV/0! error value.
    


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]