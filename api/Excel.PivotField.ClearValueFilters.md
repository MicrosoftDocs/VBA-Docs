---
title: PivotField.ClearValueFilters method (Excel)
keywords: vbaxl10.chm240155
f1_keywords:
- vbaxl10.chm240155
ms.prod: excel
api_name:
- Excel.PivotField.ClearValueFilters
ms.assetid: 8a1e12a6-0f21-bc5d-3c63-b67f534172b6
ms.date: 05/04/2019
localization_priority: Normal
---


# PivotField.ClearValueFilters method (Excel)

Calling this method deletes all value filters in the **[PivotFilters](excel.pivotfilters.md)** collection of the PivotField.


## Syntax

_expression_.**ClearValueFilters**

_expression_ A variable that represents a **[PivotField](Excel.PivotField.md)** object.


## Remarks

The following list contains the different value filter types that will be deleted by this method:

- **xlTopCount**
- **xlBottomCount**
- **xlTopPercent**
- **xlBottomPercent**
- **xlTopSum**
- **xlBottomSum**
- **xlValueEquals**
- **xlValueDoesNotEqual**
- **xlValueIsGreaterThan**
- **xlValueIsGreaterThanOrEqualTo**
- **xlValueIsLessThan**
- **xlValueIsLessThanOrEqualTo**
- **xlValueIsBetween**
- **xlValueIsNotBetween**



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]