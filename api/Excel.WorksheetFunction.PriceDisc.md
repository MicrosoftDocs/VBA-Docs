---
title: WorksheetFunction.PriceDisc method (Excel)
keywords: vbaxl10.chm137313
f1_keywords:
- vbaxl10.chm137313
ms.prod: excel
api_name:
- Excel.WorksheetFunction.PriceDisc
ms.assetid: 973cf0b8-4bb5-bc0b-f32b-59d46b8474ef
ms.date: 06/08/2017
localization_priority: Normal
---


# WorksheetFunction.PriceDisc method (Excel)

Returns the price per $100 face value of a discounted security.


## Syntax

_expression_. `PriceDisc`( `_Arg1_` , `_Arg2_` , `_Arg3_` , `_Arg4_` , `_Arg5_` )

_expression_ A variable that represents a [WorksheetFunction](./Excel.WorksheetFunction.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Variant**|Settlement - the security's settlement date. The security settlement date is the date after the issue date when the security is traded to the buyer.|
| _Arg2_|Required| **Variant**|Maturity - the security's maturity date. The maturity date is the date when the security expires.|
| _Arg3_|Required| **Variant**|Discount - the security's discount rate.|
| _Arg4_|Required| **Variant**|Redemption - the security's redemption value per $100 face value.|
| _Arg5_|Optional| **Variant**|Basis - the type of day count basis to use.|

## Return value

Double


## Remarks


 **Important**  Dates should be entered by using the DATE function, or as results of other formulas or functions. For example, use DATE(2008,5,23) for the 23rd day of May, 2008. Problems can occur if dates are entered as text.



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
    
- Settlement, maturity, and basis are truncated to integers.
    
- If settlement or maturity is not a valid date, PRICEDISC returns the #VALUE! error value.
    
- If discount ? 0 or if redemption ? 0, PRICEDISC returns the #NUM! error value.
    
- If basis < 0 or if basis > 4, PRICEDISC returns the #NUM! error value.
    
- If settlement ? maturity, PRICEDISC returns the #NUM! error value.
    
- PRICEDISC is calculated as follows:
![Formula](../images/awfprdc_ZA06051234.gif)where: B = number of days in year, depending on year basis. DSM = number of days from settlement to maturity. 
    

## See also


[WorksheetFunction Object](Excel.WorksheetFunction.md)

