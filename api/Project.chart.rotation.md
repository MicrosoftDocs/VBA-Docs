---
title: Chart.Rotation property (Project)
ms.prod: project-server
ms.assetid: a6281031-fb66-6b79-47c2-d6708c997f32
ms.date: 06/08/2017
localization_priority: Normal
---


# Chart.Rotation property (Project)
Gets or sets the rotation of the 3D chart view (the rotation of the plot area around the z-axis), in degrees. Read/write  **Variant**.

## Syntax

_expression_.**Rotation**

_expression_ A variable that represents a **[Chart](Project.Chart.md)** object.


## Remarks

The value of the **Rotation** property must be from 0 to 360, except for 3D bar charts, where the value must be from 0 to 44. The default value is 20.

Rotations are rounded to the nearest integer.


## Example

The following example sets the rotation of the chart to 45 degrees. The example should be run on a 3D chart.


```vb
Sub SetRotation()
    Dim chartShape As Shape
    Dim reportName As String
    
    reportName = "Simple 3D chart"
    Set chartShape = ActiveProject.Reports(reportName).Shapes(1)
    
    chartShape.Chart.Rotation = 45
End Sub
```


## Property value

 **VARIANT**


## See also


[Chart Object](Project.chart.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]