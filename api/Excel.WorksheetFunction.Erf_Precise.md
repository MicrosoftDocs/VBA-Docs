---
title: WorksheetFunction.Erf_Precise Method (Excel)
keywords: vbaxl10.chm137416
f1_keywords:
- vbaxl10.chm137416
ms.prod: excel
api_name:
- Excel.WorksheetFunction.Erf_Precise
ms.assetid: 1a34f60c-b5e9-f18f-2d0e-4ebe770edd59
ms.date: 06/08/2017
---


# WorksheetFunction.Erf_Precise Method (Excel)

Returns the error function integrated between zero and lower_limit.


## Syntax

 _expression_. `Erf_Precise`( `_Arg1_` )

 _expression_ A variable that represents a '[WorksheetFunction](Excel.WorksheetFunction.md)' object.


### Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Variant**|Lower_limit - the lower bound for integrating ERF.|

### Return value

Double


## Remarks


- If lower_limit is nonnumeric,  **Erf_Precise** generates an error.
    
- If lower_limit is negative,  **Erf_Precise** generates an error.
![Formula](../images/awferf1_ZA06051136.gif)


    

 **Note**  If you previously used the  **[Erf](Excel.WorksheetFunction.Erf.md)** method, which provides an optional upper_limit parameter, using the **Erf_Precise** method is equivalent to calling Erf(lower_limit) or Erf(0, upper_limit)


## See also


[WorksheetFunction Object](Excel.WorksheetFunction.md)

