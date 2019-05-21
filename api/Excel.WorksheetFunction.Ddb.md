---
title: WorksheetFunction.Ddb method (Excel)
keywords: vbaxl10.chm137135
f1_keywords:
- vbaxl10.chm137135
ms.prod: excel
api_name:
- Excel.WorksheetFunction.Ddb
ms.assetid: 7514f3b3-ca21-ec3f-28c5-f34281fc1a1f
ms.date: 05/22/2019
localization_priority: Normal
---


# WorksheetFunction.Ddb method (Excel)

Returns the depreciation of an asset for a specified period by using the double-declining balance method or some other method that you specify.


## Syntax

_expression_.**Ddb** (_Arg1_, _Arg2_, _Arg3_, _Arg4_, _Arg5_)

_expression_ A variable that represents a **[WorksheetFunction](Excel.WorksheetFunction.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Double**|Cost - the initial cost of the asset.|
| _Arg2_|Required| **Double**|Salvage - the value at the end of the depreciation (sometimes called the salvage value of the asset). This value can be 0.|
| _Arg3_|Required| **Double**|Life - the number of periods over which the asset is being depreciated (sometimes called the useful life of the asset).|
| _Arg4_|Required| **Double**|Period - the period for which you want to calculate the depreciation. Period must use the same units as life.|
| _Arg5_|Optional| **Variant**|Factor - the rate at which the balance declines. If factor is omitted, it is assumed to be 2 (the double-declining balance method).|

## Return value

**Double**


## Remarks

> [!IMPORTANT] 
> All five arguments must be positive numbers.

The double-declining balance method computes depreciation at an accelerated rate. Depreciation is highest in the first period and decreases in successive periods. **Ddb** uses the following formula to calculate depreciation for a period:

> `Min( (cost - total depreciation from prior periods) * (factor/life), (cost - salvage - total depreciation from prior periods) )`
    
Change factor if you do not want to use the double-declining balance method.
    
Use the **[VDB](excel.worksheetfunction.vdb.md)** function if you want to switch to the straight-line depreciation method when depreciation is greater than the declining balance calculation.
    



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]