---
title: WorksheetFunction.Db method (Excel)
keywords: vbaxl10.chm137171
f1_keywords:
- vbaxl10.chm137171
ms.prod: excel
api_name:
- Excel.WorksheetFunction.Db
ms.assetid: 09c56126-ab90-1bb2-44e9-3d5346ddc72d
ms.date: 05/22/2019
localization_priority: Normal
---


# WorksheetFunction.Db method (Excel)

Returns the depreciation of an asset for a specified period using the fixed-declining balance method.


## Syntax

_expression_.**Db** (_Arg1_, _Arg2_, _Arg3_, _Arg4_, _Arg5_)

_expression_ A variable that represents a **[WorksheetFunction](Excel.WorksheetFunction.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Double**|Cost - the initial cost of the asset.|
| _Arg2_|Required| **Double**|Salvage - the value at the end of the depreciation (sometimes called the salvage value of the asset).|
| _Arg3_|Required| **Double**|Life - the number of periods over which the asset is being depreciated (sometimes called the useful life of the asset).|
| _Arg4_|Required| **Double**|Period - the period for which you want to calculate the depreciation. Period must use the same units as life.|
| _Arg5_|Optional| **Variant**|Month - the number of months in the first year. If month is omitted, it is assumed to be 12.|

## Return value

**Double**


## Remarks

The fixed-declining balance method computes depreciation at a fixed rate. **Db** uses the following formulas to calculate depreciation for a period: 

> `(cost - total depreciation from prior periods) * rate` &nbsp; where &nbsp; `rate = 1 - ((salvage / cost) ^ (1 / life))`, rounded to three decimal places 
    
Depreciation for the first and last periods is a special case. For the first period, **Db** uses this formula: 

> `cost * rate * month / 12`
    
For the last period, **Db** uses this formula: 

> `((cost - total depreciation from prior periods) * rate * (12 - month)) / 12`
    



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]