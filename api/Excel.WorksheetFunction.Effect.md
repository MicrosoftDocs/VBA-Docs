---
title: WorksheetFunction.Effect method (Excel)
keywords: vbaxl10.chm137322
f1_keywords:
- vbaxl10.chm137322
ms.prod: excel
api_name:
- Excel.WorksheetFunction.Effect
ms.assetid: cbd5be5b-a1ee-addf-f0d9-01c4e4e0273b
ms.date: 05/22/2019
localization_priority: Normal
---


# WorksheetFunction.Effect method (Excel)

Returns the effective annual interest rate, given the nominal annual interest rate and the number of compounding periods per year.


## Syntax

_expression_.**Effect** (_Arg1_, _Arg2_)

_expression_ A variable that represents a **[WorksheetFunction](Excel.WorksheetFunction.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Variant**|Nominal_rate - the nominal interest rate.|
| _Arg2_|Required| **Variant**|Npery - the number of compounding periods per year.|

## Return value

**Double**


## Remarks

Npery is truncated to an integer.
    
If either argument is nonnumeric, **Effect** returns the #VALUE! error value.
    
If nominal_rate ≤ 0 or if npery < 1, **Effect** returns the #NUM! error value.
    
**Effect** is calculated as follows:

>  ![Formula](../images/awfefect_ZA06051135.gif)


    


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]