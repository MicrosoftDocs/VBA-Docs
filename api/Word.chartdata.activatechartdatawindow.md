---
title: ChartData.ActivateChartDataWindow method (Word)
keywords: vbawd10.chm190382084
f1_keywords:
- vbawd10.chm190382084
ms.prod: word
ms.assetid: dd84d89c-4c6f-27be-5e70-7ff209981cd1
ms.date: 06/08/2017
localization_priority: Normal
---


# ChartData.ActivateChartDataWindow method (Word)

Opens a Excel data grid window that contains the full source data for the specified chart.


## Syntax

_expression_. `ActivateChartDataWindow`

_expression_ A variable that represents a [ChartData](./Word.ChartData.md) object.


## Return value

 **VOID**


## Remarks

If the data grid window is already open, this method has no effect.

The  **ActivateChartDataWindow** method differs from the [ChartData.Activate](Word.ChartData.Activate.md) method in that the former opens the chart in an Excel window within Word, with the Excel ribbon unavailable, whereas the latter opens a full version of Excel, with the ribbon available.


## Example

The following example shows how to activate the chart data window for the chart that is at the first index position in the collection of shapes in the active document.


```vb

Public Sub ActivateChartDataWindow_Example()

    ThisDocument.Shapes(1).Chart.ChartData.ActivateChartDataWindow

End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]