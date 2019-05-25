---
title: WorksheetFunction.Slope method (Excel)
keywords: vbaxl10.chm137219
f1_keywords:
- vbaxl10.chm137219
ms.prod: excel
api_name:
- Excel.WorksheetFunction.Slope
ms.assetid: 26191331-d4eb-d054-b124-c57ebf4fef13
ms.date: 05/25/2019
localization_priority: Normal
---


# WorksheetFunction.Slope method (Excel)

Returns the slope of the linear regression line through data points in known_y's and known_x's. The slope is the vertical distance divided by the horizontal distance between any two points on the line, which is the rate of change along the regression line.

## Syntax

_expression_.**Slope** (_Arg1_, _Arg2_)

_expression_ A variable that represents a **[WorksheetFunction](Excel.WorksheetFunction.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Variant**|Known_y's - an array or cell range of numeric dependent data points.|
| _Arg2_|Required| **Variant**|Known_x's - the set of independent data points.|

## Return value

**Double**


## Remarks

The arguments must be either numbers or names, arrays, or references that contain numbers.
    
If an array or reference argument contains text, logical values, or empty cells, those values are ignored; however, cells with the value zero are included.
    
If known_y's and known_x's are empty or have a different number of data points, **Slope** returns the #N/A error value.
    
The equation for the slope of the regression line is as follows, where x and y are the sample means AVERAGE(known_x's) and AVERAGE(known_y's):

> ![Formula](../images/awfintc2_ZA06051175.gif)
    
The underlying algorithm used in the **Slope** and **Intercept** functions is different than the underlying algorithm used in the **LinEst** function. The difference between these algorithms can lead to different results when data is undetermined and collinear. For example, if the data points of the known_y's argument are 0 and the data points of the known_x's argument are 1: 
    
- **Slope** and **Intercept** return a #DIV/0! error. The **Slope** and **Intercept** algorithm is designed to look for one and only one answer, and in this case, there can be more than one answer.
    
- **LinEst** returns a value of 0. The **LinEst** algorithm is designed to return reasonable results for collinear data, and in this case, at least one answer can be found.
    

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]