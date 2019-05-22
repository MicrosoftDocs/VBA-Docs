---
title: WorksheetFunction.Intercept method (Excel)
keywords: vbaxl10.chm137215
f1_keywords:
- vbaxl10.chm137215
ms.prod: excel
api_name:
- Excel.WorksheetFunction.Intercept
ms.assetid: 8fa9d911-24af-6e1c-0c0b-b42b18e75e10
ms.date: 05/23/2019
localization_priority: Normal
---


# WorksheetFunction.Intercept method (Excel)

Calculates the point at which a line will intersect the y-axis by using existing x-values and y-values. The intercept point is based on a best-fit regression line plotted through the known x-values and known y-values. 

Use the **Intercept** function when you want to determine the value of the dependent variable when the independent variable is 0 (zero). For example, you can use the **Intercept** function to predict a metal's electrical resistance at 0°C when your data points were taken at room temperature and higher.


## Syntax

_expression_.**Intercept** (_Arg1_, _Arg2_)

_expression_ A variable that represents a **[WorksheetFunction](Excel.WorksheetFunction.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Variant**|Known_y's - the dependent set of observations or data.|
| _Arg2_|Required| **Variant**|Known_x's - the independent set of observations or data.|

## Return value

**Double**


## Remarks

The arguments should be either numbers or names, arrays, or references that contain numbers.
    
If an array or reference argument contains text, logical values, or empty cells, those values are ignored; however, cells with the value zero are included.
    
If known_y's and known_x's contain a different number of data points or contain no data points, **Intercept** returns the #N/A error value.
    
The equation for the intercept of the regression line, a, is:

> ![Formula](../images/awfintc1_ZA06051174.gif)

where the slope, b, is calculated as: 

> ![Formula](../images/awfintc2_ZA06051175.gif)

and where x and y are the sample means AVERAGE(known_x's) and AVERAGE(known_y's). 
    
The underlying algorithm used in the **Intercept** and **[Slope](excel.worksheetfunction.slope.md)** functions is different than the underlying algorithm used in the **[LinEst](excel.worksheetfunction.linest.md)** function. The difference between these algorithms can lead to different results when data is undetermined and collinear. For example, if the data points of the known_y's argument are 0 and the data points of the known_x's argument are 1: 
    
- **Intercept** and **Slope** return a #DIV/0! error. The **Intercept** and **Slope** algorithm is designed to look for one and only one answer, and in this case there can be more than one answer.
    
- **LinEst** returns a value of 0. The **LinEst** algorithm is designed to return reasonable results for collinear data, and in this case at least one answer can be found.
    



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]