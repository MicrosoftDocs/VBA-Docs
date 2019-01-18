---
title: WorksheetFunction.ChiSq_Inv method (Excel)
keywords: vbaxl10.chm137400
f1_keywords:
- vbaxl10.chm137400
ms.prod: excel
api_name:
- Excel.WorksheetFunction.ChiSq_Inv
ms.assetid: 1fa20ff7-e7e9-fe08-fd0f-d109af8037d1
ms.date: 06/08/2017
localization_priority: Normal
---


# WorksheetFunction.ChiSq_Inv method (Excel)

Returns the inverse of the left-tailed probability of the chi-squared distribution.


## Syntax

_expression_. `ChiSq_Inv`( `_Arg1_` , `_Arg2_` )

_expression_ A variable that represents a '[WorksheetFunction](Excel.WorksheetFunction.md)' object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Double**|Probability - A probability associated with the chi-squared distribution.|
| _Arg2_|Required| **Double**|Deg_freedom - The number of degrees of freedom.|

## Return value

Double


## Remarks




- If any argument is nonnumeric, CHISQ_INV returns the #VALUE! error value. 
    
- If probability < 0 or probability > 1, CHISQ_INV returns the #NUM! error value. 
    
- If deg_freedom is not an integer, it is truncated. 
    



## See also


[WorksheetFunction Object](Excel.WorksheetFunction.md)

