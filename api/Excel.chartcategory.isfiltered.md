---
title: ChartCategory.IsFiltered property (Excel)
keywords: vbaxl10.chm946075
f1_keywords:
- vbaxl10.chm946075
ms.prod: excel
ms.assetid: 1bf69302-7c43-3db2-1f11-6c0a72f3886e
ms.date: 04/20/2019
localization_priority: Normal
---


# ChartCategory.IsFiltered property (Excel)

Returns **True** when the user filters out a series. Read-only **Boolean**.


## Syntax

_expression_.**IsFiltered**

_expression_ A variable that represents a **[ChartCategory](Excel.chartcategory.md)** object.


## Remarks

When a series is transferred out of its parent **SeriesCollection**, that series still remains in its parent **FullSeriesCollection**. When a user filters the series back in, it is inserted back in its original place in the **SeriesCollection**.


## Property value

**BOOL**



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]