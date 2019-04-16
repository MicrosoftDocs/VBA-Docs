---
title: Chart.ChartStyle property (Project)
ms.prod: project-server
ms.assetid: e90f17dd-b9a8-4da1-d66a-2940e47953b5
ms.date: 06/08/2017
localization_priority: Normal
---


# Chart.ChartStyle property (Project)
Gets or sets the chart style for a chart. Read/write  **Variant**.

## Syntax

_expression_. `ChartStyle`

_expression_ A variable that represents a **[Chart](Project.Chart.md)** object.


## Remarks

You can use a number from 1 to 48 to set the chart style.


## Example

To use the following  **CycleThroughStyles** method, make a chart active, and then set a breakpoint in the **For … Next** loop to observe the chart styles.


```vb
Sub CycleThroughStyles()
    Dim chartShape As Shape
    Dim reportName As String
    Dim i As Integer
    
    reportName = "Simple scalar chart"
    Set chartShape = ActiveProject.Reports(reportName).Shapes(1)
    
    For i = 1 To 48
        chartShape.Chart.ChartStyle = i
    Next i
End Sub
```


## Property value

 **VARIANT**


## See also


[Chart Object](Project.chart.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]