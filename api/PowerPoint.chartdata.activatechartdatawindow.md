---
title: ChartData.ActivateChartDataWindow method (PowerPoint)
keywords: vbapp10.chm689005
f1_keywords:
- vbapp10.chm689005
ms.assetid: 3364ab9c-ed34-5970-6318-95a694a55354
ms.date: 06/08/2017
ms.prod: powerpoint
localization_priority: Normal
---


# ChartData.ActivateChartDataWindow method (PowerPoint)

Opens a Excel data grid window that contains the full source data for the specified chart.


## Syntax

_expression_. `ActivateChartDataWindow`

_expression_ A variable that represents a [ChartData](PowerPoint.ChartData.md) object.


## Return value

 **VOID**


## Remarks

If the data grid window is already open, this method has no effect.

The  **ActivateChartDataWindow** method differs from the [ChartData.Activate](PowerPoint.ChartData.Activate.md) method in that the former opens the chart in an Excel window within Word, with the Excel ribbon unavailable, whereas the latter opens a full version of Excel, with the ribbon available.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]