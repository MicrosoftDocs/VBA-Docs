---
title: PivotField.ClearLabelFilters method (Excel)
keywords: vbaxl10.chm240156
f1_keywords:
- vbaxl10.chm240156
ms.prod: excel
api_name:
- Excel.PivotField.ClearLabelFilters
ms.assetid: 48b8f6be-b4c0-26c6-2550-63729fd6a918
ms.date: 05/04/2019
localization_priority: Normal
---


# PivotField.ClearLabelFilters method (Excel)

This method deletes all label filters or all date filters in the **[PivotFilters](excel.pivotfilters.md)** collection of the PivotField.


## Syntax

_expression_.**ClearLabelFilters**

_expression_ A variable that represents a **[PivotField](Excel.PivotField.md)** object.


## Remarks

The following list contains the different label filter types that will be deleted by this method:

- **xlCaptionEquals**
- **xlCaptionDoesNotEqual**
- **xlCaptionIsGreaterThan**
- **xlCaptionIsGreaterThanOrEqualTo**
- **xlCaptionIsLessThan**
- **xlCaptionIsLessThanOrEqualTo**
- **xlCaptionBeginsWith**
- **xlCaptionDoesNotBeginWith**
- **xlCaptionEndsWith**
- **xlCaptionDoesNotEndWith**
- **xlCaptionContains**
- **xlCaptionDoesNotContain**
- **xlCaptionIsBetween**
- **xlCaptionIsNotBetween**


The following list contains the different date filter types that will be deleted by this method:

- **xlSpecificDate**
- **xlNotSpecificDate**
- **xlBefore**
- **xlBeforeOrEqualTo**
- **xlAfter**
- **xlAfterOrEqualTo**
- **xlDateBetween**
- **xlDateNotBetween**
- **xlDateToday**
- **xlDateYesterday**
- **xlDateTomorrow**
- **xlDateThisWeek**
- **xlDateLastWeek**
- **xlDateNextWeek**
- **xlDateThisMonth**
- **xlDateLastMonth**
- **xlDateNextMonth**
- **xlDateThisQuarter**
- **xlDateLastQuarter**
- **xlDateNextQuarter**
- **xlDateThisYear**
- **xlDateLastYear**
- **xlDateNextYear**
- **xlYearToDate**
- **xlAllDatesInPeriodQuarter1**
- **xlAllDatesInPeriodQuarter2**
- **xlAllDatesInPeriodQuarter3**
- **xlAllDatesInPeriodQuarter4**
- **xlAllDatesInPeriodJanuary**
- **xlAllDatesInPeriodFebruary**
- **xlAllDatesInPeriodMarch**
- **xlAllDatesInPeriodApril**
- **xlAllDatesInPeriodMay**
- **xlAllDatesInPeriodJune**
- **xlAllDatesInPeriodJuly**
- **xlAllDatesInPeriodAugust**
- **xlAllDatesInPeriodSeptember**
- **xlAllDatesInPeriodOctober**
- **xlAllDatesInPeriodNovember**
- **xlAllDatesInPeriodDecember**



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]