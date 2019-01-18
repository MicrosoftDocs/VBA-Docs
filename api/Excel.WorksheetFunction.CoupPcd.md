---
title: WorksheetFunction.CoupPcd method (Excel)
keywords: vbaxl10.chm137333
f1_keywords:
- vbaxl10.chm137333
ms.prod: excel
api_name:
- Excel.WorksheetFunction.CoupPcd
ms.assetid: 68b08847-2729-3975-f488-241c20b398a5
ms.date: 06/08/2017
localization_priority: Normal
---


# WorksheetFunction.CoupPcd method (Excel)




## Syntax

_expression_. `CoupPcd`( `_Arg1_` , `_Arg2_` , `_Arg3_` , `_Arg4_` )

_expression_ A variable that represents a [WorksheetFunction](./Excel.WorksheetFunction.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Variant**|The security's settlement date. The security settlement date is the date after the issue date when the security is traded to the buyer.|
| _Arg2_|Required| **Variant**|The security's maturity date. The maturity date is the date when the security expires.|
| _Arg3_|Required| **Variant**|The number of coupon payments per year. For annual payments, frequency = 1; for semiannual, frequency = 2; for quarterly, frequency = 4.|
| _Arg4_|Optional| **Variant**|The type of day count basis to use.|

## Return value

Double


## Remarks


 **Important**  Dates should be entered by using the DATE function, or as results of other formulas or functions. For example, use DATE(2008,5,23) for the 23rd day of May, 2008. Problems can occur if dates are entered as text.

The following table contains the values used in  _Arg4_ .



|**Basis**|**Day count basis**|
|:-----|:-----|
|0 or omitted|US (NASD) 30/360|
|1|Actual/actual|
|2|Actual/360|
|3|Actual/365|
|4|European 30/360|

- Microsoft Excel stores dates as sequential serial numbers so they can be used in calculations. By default, January 1, 1900 is serial number 1, and January 1, 2008 is serial number 39448 because it is 39,448 days after January 1, 1900. Microsoft Excel for the Macintosh uses a different date system as its default.
    
     **Note**  Visual Basic for Applications (VBA) calculates serial dates differently than Excel. In VBA, serial number 1 is December 31, 1899, rather than January 1, 1900. 
- The settlement date is the date a buyer purchases a coupon, such as a bond. The maturity date is the date when a coupon expires. For example, suppose a 30-year bond is issued on January 1, 2008, and is purchased by a buyer six months later. The issue date would be January 1, 2008, the settlement date would be July 1, 2008, and the maturity date would be January 1, 2038, which is 30 years after the January 1, 2008, issue date.
    
- All arguments are truncated to integers.
    
- If settlement or maturity is not a valid date, CoupPcd generates an error.
    
- If frequency is any number other than 1, 2, or 4, CoupPcd generates an error.
    
- If basis < 0 or if basis > 4, CoupPcd generates an error.
    
- If settlement ? maturity, CoupPcd generates an error.
    

## See also


[WorksheetFunction Object](Excel.WorksheetFunction.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]