---
title: Workbook.NewChart event (Excel)
keywords: vbaxl10.chm503108
f1_keywords:
- vbaxl10.chm503108
ms.prod: excel
api_name:
- Excel.Workbook.NewChart
ms.assetid: 76e7f325-9244-fd8c-b38d-063f0193a5e9
ms.date: 05/29/2019
localization_priority: Normal
---


# Workbook.NewChart event (Excel)

Occurs when a new chart is created in the workbook.


## Syntax

_expression_.**NewChart** (_Ch_)

_expression_ A variable that returns a **[Workbook](Excel.Workbook.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Ch_|Required| **[Chart](Excel.Chart(object).md)**|The new chart.|

## Return value

**Nothing**


## Remarks

The **NewChart** event occurs whenever a new chart is inserted or pasted on a worksheet, a chart sheet, or other sheet types. If multiple charts are inserted or pasted, the event will occur for each chart in the order they are inserted or pasted. If a chart object or chart sheet is moved from one location to another, the event will not occur. However, if the chart is moved between a chart object and a chart sheet, the event will occur because a new chart must be created.

The **NewChart** event will not occur in the following scenarios: copying or pasting a chart sheet, changing a chart type, changing a chart data source, undoing or redoing inserting or pasting a chart, and loading a workbook that contains a chart.


## Example

This example displays a message box when a new chart is added to the workbook.

```vb
Private Sub Workbook_NewChart(ByVal Ch As Chart) 
 MsgBox ("A new chart of the following chart type was added: " & Ch.ChartType) 
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]