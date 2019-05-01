---
title: PivotFilter.WholeDayFilter property (Excel)
keywords: vbaxl10.chm770086
f1_keywords:
- vbaxl10.chm770086
ms.prod: excel
ms.assetid: 4dc32caf-50de-0cd0-a3d7-b8b52deb4370
ms.date: 06/08/2017
localization_priority: Normal
---


# PivotFilter.WholeDayFilter property (Excel)

Sets or gets the filtering semantics for date filters. Read/write **Boolean**.


## Syntax

_expression_. `WholeDayFilter`

_expression_ A variable that represents a **[PivotFilter](Excel.PivotFilter.md)** object.


## Remarks

The following describes the results for previous and current property settings: 


- False (Microsoft Office2010 behavior): Any time can be specified; dates are precise points in time (midnight of the specified date). Filtering date ranges start or end at midnight.
    
- True (Microsoft Office 2013 behavior): This behavior is enforced for Timeline controls. Only whole dates are specified; dates include all times-of-day until and not including the next day at midnight.
    
For a Timeline, always returns **True**; returns a run-time error when setting this to **False**.

For a non-date filter, any access returns a run-time error.


## Property value

 **BOOL**


## See also


[PivotFilter Object](Excel.PivotFilter.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]