---
title: WorksheetFunction.EoMonth method (Excel)
keywords: vbaxl10.chm137326
f1_keywords:
- vbaxl10.chm137326
ms.prod: excel
api_name:
- Excel.WorksheetFunction.EoMonth
ms.assetid: 46ffb33b-2992-88d4-59ed-5c0660fbbf5d
ms.date: 05/22/2019
localization_priority: Normal
---


# WorksheetFunction.EoMonth method (Excel)

Returns the serial number for the last day of the month that is the indicated number of months before or after start_date. Use **EoMonth** to calculate maturity dates or due dates that fall on the last day of the month.


## Syntax

_expression_.**EoMonth** (_Arg1_, _Arg2_)

_expression_ A variable that represents a **[WorksheetFunction](Excel.WorksheetFunction.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Variant**|Start_date - a date that represents the starting date. Dates should be entered by using the DATE function, or as results of other formulas or functions. For example, use DATE(2008,5,23) for the 23rd day of May, 2008. Problems can occur if dates are entered as text.|
| _Arg2_|Required| **Variant**|Months - the number of months before or after start_date. A positive value for months yields a future date; a negative value yields a past date.|

## Return value

**Double**


## Remarks

Microsoft Excel stores dates as sequential serial numbers so they can be used in calculations. By default, January 1, 1900 is serial number 1, and January 1, 2008 is serial number 39448 because it is 39,448 days after January 1, 1900. Microsoft Excel for the Macintosh uses a different date system as its default.
    
> [!NOTE] 
> Visual Basic for Applications (VBA) calculates serial dates differently than Excel. In VBA, serial number 1 is December 31, 1899, rather than January 1, 1900. 

If start_date is not a valid date, **EoMonth** returns the #NUM! error value.
    
If start_date plus months yields an invalid date, **EoMonth** returns the #NUM! error value.
    


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
