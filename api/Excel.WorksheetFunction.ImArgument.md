---
title: WorksheetFunction.ImArgument method (Excel)
keywords: vbaxl10.chm137284
f1_keywords:
- vbaxl10.chm137284
ms.prod: excel
api_name:
- Excel.WorksheetFunction.ImArgument
ms.assetid: ac1e721a-edfe-0287-afa1-509f5c437cd8
ms.date: 06/08/2017
localization_priority: Normal
---


# WorksheetFunction.ImArgument method (Excel)

Returns the argument 
![Formula](../images/theta_ZA06052070.gif) (theta), an angle expressed in radians, such that:
![Formula](../images/awfimar1_ZA06051153.gif)




## Syntax

_expression_. `ImArgument`( `_Arg1_` )

_expression_ A variable that represents a [WorksheetFunction](./Excel.WorksheetFunction.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Variant**|Inumber is a complex number for which you want the argument theta.|

## Return value

String


## Remarks




- Use COMPLEX to convert real and imaginary coefficients into a complex number.
    
- IMARGUMENT is calculated as follows:
![Formula](../images/awfimar2_ZA06051154.gif)where: 
![Formula](../images/awfimar3_ZA06051155.gif) and z = x + yi
    

## See also


[WorksheetFunction Object](Excel.WorksheetFunction.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]