---
title: WorksheetFunction.F_Inv method (Excel)
keywords: vbaxl10.chm137404
f1_keywords:
- vbaxl10.chm137404
ms.prod: excel
api_name:
- Excel.WorksheetFunction.F_Inv
ms.assetid: c24c12b0-9c0b-076c-4488-947ec94f8dd0
ms.date: 05/22/2019
localization_priority: Normal
---


# WorksheetFunction.F_Inv method (Excel)

Returns the inverse of the F probability distribution.


## Syntax

_expression_.**F_Inv** (_Arg1_, _Arg2_, _Arg3_)

_expression_ A variable that represents a **[WorksheetFunction](Excel.WorksheetFunction.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Double**|Probability - A probability associated with the F cumulative distribution.|
| _Arg2_|Required| **Double**|Deg_freedom1 - The numerator degrees of freedom.|
| _Arg3_|Required| **Double**|Deg_freedom2 - The denominator degrees of freedom.|

## Return value

**Double**


## Remarks

If any argument is nonnumeric, **F_Inv** returns the #VALUE! error value. 
    
If probability < 0 or probability > 1, **F_Inv** returns the #NUM! error value. 
    
If deg_freedom1 or deg_freedom2 is not an integer, it is truncated. 
    
If deg_freedom1 < 1, or deg_freedom2 < 1, **F_Inv** returns the #NUM! error value. 
    



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]