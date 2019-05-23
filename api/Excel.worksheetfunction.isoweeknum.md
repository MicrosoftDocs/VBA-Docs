---
title: WorksheetFunction.IsoWeekNum method (Excel)
keywords: vbaxl10.chm137457
f1_keywords:
- vbaxl10.chm137457
ms.prod: excel
ms.assetid: 8b643312-d9b9-c509-ca9f-c3d960ba012c
ms.date: 05/23/2019
localization_priority: Normal
---


# WorksheetFunction.IsoWeekNum method (Excel)

Returns the ISO week number of the year for a given date. 


## Syntax

_expression_.**IsoWeekNum** (_Arg1_, _Arg2_)

_expression_ A variable that represents a **[WorksheetFunction](Excel.WorksheetFunction.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required|**Double**|Date-time code used by Microsoft Excel for date and time calculation.|
| _Arg2_|Optional|**Variant**|This argument is not available in the function.|

## Return value

**Double**


## Remarks

Returns the ordinal number of the [ISO 8601](https://en.wikipedia.org/wiki/ISO_8601) calendar week in the year for the given date. ISO 8601 defines the calendar week as a time interval of seven calendar days starting with a Monday, and the first calendar week of a year as the one that includes the first Thursday of that year.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]