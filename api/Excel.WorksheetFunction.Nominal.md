---
title: WorksheetFunction.Nominal method (Excel)
keywords: vbaxl10.chm137321
f1_keywords:
- vbaxl10.chm137321
api_name:
- Excel.WorksheetFunction.Nominal
ms.assetid: 4ba61f10-233b-400b-76e1-90147fd7f503
ms.date: 05/24/2019
ms.localizationpriority: medium
---


# WorksheetFunction.Nominal method (Excel)

Returns the nominal annual interest rate, given the effective rate and the number of compounding periods per year.


## Syntax

_expression_.**Nominal** (_Arg1_, _Arg2_)

_expression_ A variable that represents a **[WorksheetFunction](Excel.WorksheetFunction.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Variant**|Effect_rate - the effective interest rate.|
| _Arg2_|Required| **Variant**|Npery - the number of compounding periods per year.|

## Return value

**Double**


## Remarks

Npery is truncated to an integer.
    
If either argument is nonnumeric, **Nominal** returns the #VALUE! error value.
    
If effect_rate ≤ 0 or if npery < 1, **Nominal** returns the #NUM! error value.
    
**Nominal** is related to **[Effect](excel.worksheetfunction.effect.md)** as shown in the following equation:

> ![Formula](../images/awfnomnl_ZA06051211.gif)




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]