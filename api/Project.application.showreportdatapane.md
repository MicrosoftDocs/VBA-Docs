---
title: Application.ShowReportDataPane method (Project)
keywords: vbapj.chm152
f1_keywords:
- vbapj.chm152
ms.prod: project-server
ms.assetid: 7f0e991a-df7c-9534-45de-50d3839fbac7
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.ShowReportDataPane method (Project)
Shows or hides the report data pane, when a chart shape or table shape is selected in a report.

## Syntax

_expression_. `ShowReportDataPane` _(Show)_

_expression_ A variable that represents an **[Application](Project.Application.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Show_|Optional|**Boolean**|**True** to show the report data pane; **False** to hide the data pane. When the _Show_ parameter is missing, **ShowReportDataPane** toggles the report data pane between visible and not visible.|

## Return value

 **Boolean**

 **True** if the **ShowReportDataPane** method is successful; otherwise, **False**.


## Remarks

When a chart shape or table shape is selected in a report, the  **ShowReportDataPane** method can show or hide the **Field List** data pane for the chart or table. The method corresponds to the **Show Field List** command or **Hide Field List** command in the option menu when you right-click a chart or a table.

If a chart or table is not selected, the  **ShowReportDataPane** method displays a dialog box with run-time error 1100, "The method is not available in this situation." For other views, such as the Gantt chart, the **ShowReportDataPane** method has no effect, but returns **True**.


## See also


[Application Object](Project.Application.md)



[ReportTable Object](Project.reporttable.md)
[Chart Object](Project.chart.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]