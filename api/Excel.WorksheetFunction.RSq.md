---
title: WorksheetFunction.RSq method (Excel)
keywords: vbaxl10.chm137217
f1_keywords:
- vbaxl10.chm137217
api_name:
- Excel.WorksheetFunction.RSq
ms.assetid: f6d9b270-ec48-1b53-fe96-b62dd37f1a56
ms.date: 05/25/2019
ms.localizationpriority: medium
---


# WorksheetFunction.RSq method (Excel)

Returns the square of the Pearson product moment correlation coefficient through data points in known_y's and known_x's. For more information, see **[Pearson](excel.worksheetfunction.pearson.md)**. The r-squared value can be interpreted as the proportion of the variance in y attributable to the variance in x.


## Syntax

_expression_.**RSq** (_Arg1_, _Arg2_)

_expression_ A variable that represents a **[WorksheetFunction](Excel.WorksheetFunction.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Variant**|Known_y's - an array or range of data points.|
| _Arg2_|Required| **Variant**|Known_x's - an array or range of data points.|

## Return value

**Double**


## Remarks

Arguments can either be numbers or names, arrays, or references that contain numbers.
    
Logical values and text representations of numbers that you type directly into the list of arguments are counted.
    
If an array or reference argument contains text, logical values, or empty cells, those values are ignored; however, cells with the value zero are included.
    
Arguments that are error values or text that cannot be translated into numbers cause errors.
    
If known_y's and known_x's are empty or have a different number of data points, **RSq** returns the #N/A error value.
    
If known_y's and known_x's contain only one data point, **RSq** returns the #DIV/0! error value.
    
The equation for the Pearson product moment correlation coefficient, r, is as follows, where x and y are the sample means AVERAGE(known_x's) and AVERAGE(known_y's). **RSq** returns r2, which is the square of this correlation coefficient.

> ![Formula](../images/awfpears_ZA06051230.gif)
    



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]