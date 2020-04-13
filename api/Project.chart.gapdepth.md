---
title: Chart.GapDepth property (Project)
ms.prod: project-server
ms.assetid: 22b3c702-6b1e-140b-13a7-04b6cd4bdab2
ms.date: 06/08/2017
localization_priority: Normal
---


# Chart.GapDepth property (Project)
Gets or sets the distance between the data series in a 3D chart, as a percentage of the marker width. Read/write  **Long**.

## Syntax

_expression_.**GapDepth**

_expression_ A variable that represents a **[Chart](Project.Chart.md)** object.


## Remarks

The value of the **GapDepth** property must be between 0 and 500. The **GapDepth** property fails on 2D charts.


## Example

The following example sets the distance between the data series in the chart to 200 percent of the marker width. The example should be run on a 3D chart.


```vb
Sub SetGapDepth()
    Dim chartShape As Shape
    Dim reportName As String
    
    reportName = "Simple 3D chart"
    Set chartShape = ActiveProject.Reports(reportName).Shapes(1)
    
    chartShape.Chart.GapDepth = 200
End Sub
```


## Property value

 **INT**


## See also


[Chart Object](Project.chart.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]