---
title: Chart.RightAngleAxes property (Project)
ms.prod: project-server
ms.assetid: 51e8cde1-53c7-90ff-b5c7-72a091461f6b
ms.date: 06/08/2017
localization_priority: Normal
---


# Chart.RightAngleAxes property (Project)
 **True** if the chart axes are at right angles, independent of chart rotation or elevation. Applies only to 3D line, 3D column, and 3D bar charts. Read/write **Boolean**.

## Syntax

_expression_.**RightAngleAxes**

_expression_ A variable that represents a **[Chart](Project.Chart.md)** object.


## Remarks

If the  **RightAngleAxes** property is **True**, the  **[Perspective](Project.chart.perspective.md)** property is ignored.


## Example

The following example sets the chart axes to intersect at right angles. The example should be run on a 3D chart.


```vb
Sub SetRightAngleAxes()
    Dim chartShape As Shape
    Dim reportName As String
    
    reportName = "Simple 3D chart"
    Set chartShape = ActiveProject.Reports(reportName).Shapes(1)
    
    chartShape.Chart.RightAngleAxes = True
End Sub
```


## Property value

 **VARIANT**


## See also


[Chart Object](Project.chart.md)
[AutoScaling Property](Project.chart.autoscaling.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]