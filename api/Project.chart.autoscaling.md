---
title: Chart.AutoScaling property (Project)
ms.prod: project-server
ms.assetid: d7e1c8f7-8a2b-0474-1b4a-28a63605e929
ms.date: 06/08/2017
localization_priority: Normal
---


# Chart.AutoScaling property (Project)
 **True** if Project scales a 3D chart so that it is closer in size to the equivalent 2D chart. Read/write **Boolean**.

## Syntax

_expression_. `AutoScaling`

_expression_ A variable that represents a **[Chart](Project.Chart.md)** object.


## Remarks

For auto-scaling to work, the  **[RightAngleAxes](Project.chart.rightangleaxes.md)** property must also be **True**. 


## Example

In the following example, the chart is the first shape in the "3D chart" report. The example automatically scales the chart. The example should be run on a 3D chart.


```vb
Sub SetChartColor()
    Dim chartShape As Shape
    
    Set chartShape = ActiveProject.Reports("3D chart").Shapes(1)
    With chartShape
        .RightAngleAxes = True
        .AutoScaling = True
    End With
End Sub
```


## Property value

 **BOOL**


## See also


[Chart Object](Project.chart.md)
[RightAngleAxes Property](Project.chart.rightangleaxes.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]