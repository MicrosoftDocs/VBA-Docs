---
title: Chart.PlotBy property (Project)
keywords: vbapj.chm131635
f1_keywords:
- vbapj.chm131635
ms.prod: project-server
ms.assetid: 10483232-929b-c040-025e-059ddf2fe915
ms.date: 06/08/2017
localization_priority: Normal
---


# Chart.PlotBy property (Project)
Gets or sets the way columns or rows are used as data series on the chart. Can be one of the following  **Office.XlRowCol** constants: **xlColumns** or **xlRows**. Read/write  **Long**.

## Syntax

_expression_.**PlotBy**

_expression_ A variable that represents a **[Chart](Project.Chart.md)** object.


## Remarks

For PivotChart reports, The  **PlotBy** property is read-only and always returns **xlColumns**.


## Property value

 **XLROWCOL**


## See also


[Chart Object](Project.chart.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]