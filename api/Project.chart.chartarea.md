---
title: Chart.ChartArea property (Project)
ms.prod: project-server
ms.assetid: 384eb030-741d-e69d-cd27-d4e414d7da8c
ms.date: 06/08/2017
localization_priority: Normal
---


# Chart.ChartArea property (Project)
Gets an  **Office.IMsoChartArea** object that represents the complete chart area for the chart. Read-only **IMsoChartArea**.

## Syntax

_expression_. `ChartArea`

_expression_ A variable that represents a **[Chart](Project.Chart.md)** object.


## Remarks

To see the  **IMsoChartArea** object in the Object Browser, show the hidden members in the **Office** library.


## Example

In the following example, the chart is the first shape in the "Simple scalar chart" report. The example sets the chart area interior color to red.


```vb
Sub SetChartAreaColor()
    Dim chartShape As Shape
    Dim i As Integer
    
    Set chartShape = ActiveProject.Reports("Simple scalar chart").Shapes(1)
    
    With chartShape.Chart.ChartArea
        .Interior.ColorIndex = 3
    End With
End Sub
```


## Property value

 **IMSOCHARTAREA**


## See also


[Chart Object](Project.chart.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]