---
title: Chart.Floor property (Project)
ms.prod: project-server
ms.assetid: ae1f3f2b-e49c-63d1-f487-5d031fea20e5
ms.date: 06/08/2017
localization_priority: Normal
---


# Chart.Floor property (Project)
Gets an **Office.IMsoFloor** object that represents the floor of a 3D chart. Read-only **IMsoFloor**.

## Syntax

_expression_.**Floor**

_expression_ A variable that represents a **[Chart](Project.Chart.md)** object.


## Remarks

The **Floor** property fails on 2D charts.


## Example

The following example sets the floor color of the chart to blue. The example should be run on a 3D chart.


```vb
Sub SetFloorColor()
    Dim chartShape As Shape
    Dim reportName As String
    
    reportName = "Simple 3D chart"
    Set chartShape = ActiveProject.Reports(reportName).Shapes(1)
    
    chartShape.Chart.Floor.Interior.ColorIndex = 5
End Sub
```


## Property value

 **IMSOFLOOR**


## See also


[Chart Object](Project.chart.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]