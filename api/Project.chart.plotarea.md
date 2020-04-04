---
title: Chart.PlotArea property (Project)
ms.prod: project-server
ms.assetid: 4d378a40-7417-1c1d-7424-9eb5cc7367c2
ms.date: 06/08/2017
localization_priority: Normal
---


# Chart.PlotArea property (Project)
Gets an **Office.IMsoPlotArea** object that represents the plot area of a chart. Read-only **IMsoPlotArea**.

## Syntax

_expression_.**PlotArea**

_expression_ A variable that represents a **[Chart](Project.Chart.md)** object.


## Example

The following example sets the inside height of the plot area 30 points greater than it was set previously.


```vb
Sub SetChartPlotAreaHeight()
    Dim chartShape As Shape
    Dim reportName As String
    Dim insideHeight As Double
    
    reportName = "Simple scalar chart"
    Set chartShape = ActiveProject.Reports(reportName).Shapes(1)
    
    insideHeight = chartShape.Chart.PlotArea.InsideHeight
    chartShape.Chart.PlotArea.InsideHeight = insideHeight + 30
End Sub
```


## Property value

 **IMSOPLOTAREA**


## See also


[Chart Object](Project.chart.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]