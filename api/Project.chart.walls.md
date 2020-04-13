---
title: Chart.Walls property (Project)
ms.prod: project-server
ms.assetid: 8404e5cb-8da2-49b4-c49a-488d67457681
ms.date: 06/08/2017
localization_priority: Normal
---


# Chart.Walls property (Project)
Gets an **Office.IMsoWalls** object that represents the walls of a 3D chart. Read-only **IMsoWalls**.

## Syntax

_expression_.**Walls**

_expression_ A variable that represents a **[Chart](Project.Chart.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _fBackWall_|Optional|**Boolean**|Default value =  **True**. The  _fBackWall_ parameter has no effect in Project.|

## Example

The following example sets the wall borders of the 3D chart to a red line that is three points wide.


```vb
Sub FormatWalls()
    Dim chartShape As Shape
    Dim reportName As String
    
    reportName = "Simple 3D chart"
    Set chartShape = ActiveProject.Reports(reportName).Shapes(1)
    
    With chartShape.Chart.Walls.Border
        .Weight = 3
        .Color = &HFF
    End With
End Sub
```


## Property value

 **IMSOWALLS**


## See also


[Chart Object](Project.chart.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]