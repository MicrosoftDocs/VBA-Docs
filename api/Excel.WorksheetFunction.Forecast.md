---
title: WorksheetFunction.Forecast method (Excel)
keywords: vbaxl10.chm137213
f1_keywords:
- vbaxl10.chm137213
ms.prod: excel
api_name:
- Excel.WorksheetFunction.Forecast
ms.assetid: a4d178b3-7d68-bfc6-0f7a-e3c6d5984af6
ms.date: 05/22/2019
localization_priority: Normal
---


# WorksheetFunction.Forecast method (Excel)

Calculates, or predicts, a future value by using existing values. The predicted value is a y-value for a given x-value. The known values are existing x-values and y-values, and the new value is predicted by using linear regression. You can use this function to predict future sales, inventory requirements, or consumer trends.

> [!NOTE] 
> This member is deprecated in Office 2016 and later versions.


## Syntax

_expression_.**Forecast** (_Arg1_, _Arg2_, _Arg3_)

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

If x is nonnumeric, **Forecast** returns the #VALUE! error value.
    
If known_y's and known_x's are empty or contain a different number of data points, **Forecast** returns the #N/A error value.
    
If the variance of known_x's equals zero, **Forecast** returns the #DIV/0! error value.
    
The equation for FORECAST is a+bx, where &nbsp; ![Formula](../images/awfintc1_ZA06051174.gif) &nbsp; and &nbsp; ![Formula](../images/awfintc2_ZA06051175.gif) &nbsp; and where x and y are the sample means AVERAGE(known_x's) and AVERAGE(known_y's). 
    



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]