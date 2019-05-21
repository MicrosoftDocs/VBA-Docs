---
title: WorksheetFunction.Ceiling_Precise method (Excel)
keywords: vbaxl10.chm137419
f1_keywords:
- vbaxl10.chm137419
ms.prod: excel
api_name:
- Excel.WorksheetFunction.Ceiling_Precise
ms.assetid: 638b4577-fd80-bd94-9a31-84fe4f3ff9d0
ms.date: 05/22/2019
localization_priority: Normal
---


# WorksheetFunction.Ceiling_Precise method (Excel)

Returns the specified number rounded to the nearest multiple of significance.


## Syntax

_expression_.**Ceiling_Precise** (_Arg1_, _Arg2_)

_expression_ A variable that represents a **[WorksheetFunction](Excel.WorksheetFunction.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Double**|Number - the value that you want to round.|
| _Arg2_|Optional| **Variant**|Significance - the multiple to which you want to round.|

## Return value

**Double**


## Remarks

For example, if you want to avoid using pennies in your prices and your product is priced at $4.42, use the formula `Ceiling(4.42,0.05)` to round prices up to the nearest nickel.
 
If the significance argument is omitted, the value 1.0 is used.

Depending on the sign of the number and significance arguments, the **Ceiling_Precise** method rounds either away from or towards zero.

|Sign (_Arg1_/_Arg2_)|Rounding|
|:-----|:-----|
|-/-|Rounds toward zero.|
|+/+|Rounds away from zero.|
|-/+|Rounds toward zero.|
|+/-|Rounds away from zero.|

If either argument is nonnumeric, **Ceiling_Precise** generates an error.
    
If number is an exact multiple of significance, no rounding occurs.
    
> [!NOTE] 
> The CEILING.PRECISE algorithm is the same as the one used for the ISO.CEILING function.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]