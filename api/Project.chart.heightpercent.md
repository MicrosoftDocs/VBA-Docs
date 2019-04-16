---
title: Chart.HeightPercent property (Project)
ms.prod: project-server
ms.assetid: cb7e3a55-eb99-b02d-2242-ebdcbd954b35
ms.date: 06/08/2017
localization_priority: Normal
---


# Chart.HeightPercent property (Project)
Gets or sets the height of a 3D chart as a percentage of the chart width. Read/write  **Long**.

## Syntax

_expression_.**HeightPercent**

_expression_ A variable that represents a **[Chart](Project.Chart.md)** object.


## Remarks

The  **HeightPercent** value can be between 5 and 500 percent.


## Example

The following example sets the height of the chart to 80 percent of its width. The example should be run on a 3D chart.


```vb
Sub SetHeightPercent()
    Dim chartShape As Shape
    Dim reportName As String
    
    reportName = "Simple 3D chart"
    Set chartShape = ActiveProject.Reports(reportName).Shapes(1)
    
    chartShape.Chart.HeightPercent = 80
End Sub
```


## Property value

 **INT**


## See also


[Chart Object](Project.chart.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]