---
title: Chart.Copy method (Project)
keywords: vbapj.chm131611
f1_keywords:
- vbapj.chm131611
ms.prod: project-server
ms.assetid: 92627648-016a-0a69-52b8-bb24b1ea22d3
ms.date: 06/08/2017
localization_priority: Normal
---


# Chart.Copy method (Project)
Copies a chart.

## Syntax

_expression_.**Copy**

_expression_ A variable that represents a **[Chart](Project.Chart.md)** object.


## Return value

 **Variant**


## Example

The following example copies the chart and then pastes the chart as a picture on the active report.


```vb
Sub CopyAndPasteChart()
    Dim chartShape As Shape
    Dim reportName As String
    Dim duplicateChart As Chart
    
    reportName = "Simple scalar chart"
    Set chartShape = ActiveProject.Reports(reportName).Shapes(1)
    
    chartShape.Chart.Copy
    Application.PasteAsPicture
End Sub
```


## See also


[Chart Object](Project.chart.md)
[CopyPicture Method](Project.chart.copypicture.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]