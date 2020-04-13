---
title: Chart.DisplayBlanksAs property (Project)
ms.prod: project-server
ms.assetid: 241fcca1-b736-799f-9f53-17751622e1e6
ms.date: 06/08/2017
localization_priority: Normal
---


# Chart.DisplayBlanksAs property (Project)
Gets or sets the way that blank cells are plotted on a chart. Can be one of the **Office.XlDisplayBlanksAs** constants. Read/write **Long**.

## Syntax

_expression_.**DisplayBlanksAs**

_expression_ A variable that represents a **[Chart](Project.Chart.md)** object.


## Example

The following example hides blank cells in the chart.


```vb
Sub HideBlankCells()
    Dim chartShape As Shape
    Dim reportName As String
    
    reportName = "Simple scalar chart"
    Set chartShape = ActiveProject.Reports(reportName).Shapes(1)
    
    chartShape.Chart.DisplayBlanksAs = Office.XlDisplayBlanksAs.xlNotPlotted
End Sub
```


## Property value

 **XLDISPLAYBLANKSAS**


## See also


[Chart Object](Project.chart.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]