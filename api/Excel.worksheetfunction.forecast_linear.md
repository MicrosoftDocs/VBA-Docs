---
title: WorksheetFunction.Forecast_Linear method (Excel)
keywords: vbaxl10.chm137471
f1_keywords:
- vbaxl10.chm137471
ms.assetid: 71b85d12-0c81-f82d-99fe-ad712f2530e5
ms.date: 05/22/2019
ms.prod: excel
localization_priority: Normal
---


# WorksheetFunction.Forecast_Linear method (Excel)

Calculates, or predicts, a future value by using existing values. The predicted value is a y-value for a given x-value. The known values are existing x-values and y-values, and the new value is predicted by using linear regression. You can use this function to predict future sales, inventory requirements, or consumer trends.


## Syntax

_expression_.**Forecast_Linear** (_Arg1_, _Arg2_, _Arg3_)

_expression_ A variable that represents a **[WorksheetFunction](Excel.WorksheetFunction.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Double**|x - the data point for which you want to predict a value.|
| _Arg2_|Required| **Variant**|known_y's - the dependent array or range of data.|
| _Arg3_|Required| **Variant**|known_x's - the independent array or range of data.|


## Return value

**Double**


## Remarks

If x is nonnumeric, **Forecast_Linear** returns the #VALUE! error value.
    
If _known_y_ and _known_x_ parameters are empty or contain a different number of data points, **Forecast_Linear** returns the #N/A error value.
    
If the variance of _known_x_ parameters equals zero, **Forecast_Linear** returns the #DIV/0! error value.
    
The equation for **Forecast_Linear** is a+bx, where &nbsp; ![Formula](../images/awfintc1_ZA06051174.gif) &nbsp; and &nbsp; ![Formula](../images/awfintc2_ZA06051175.gif) &nbsp; and where x and y are the sample means AVERAGE(all _known_x_) and AVERAGE(all _known_y_).
    

## Example

```vb
Dim instance As WorksheetFunction
Dim Arg1 As Double
Dim Arg2 As Object
Dim Arg3 As Object
Dim returnValue As Double

returnValue = instance.Forecast_Linear(Arg1, Arg2, Arg3)

```


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]