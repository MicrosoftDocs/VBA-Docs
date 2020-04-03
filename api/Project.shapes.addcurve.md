---
title: Shapes.AddCurve method (Project)
ms.prod: project-server
ms.assetid: 16ea0f55-268a-b224-cc94-3d7e74de6265
ms.date: 06/08/2017
localization_priority: Normal
---


# Shapes.AddCurve method (Project)
Adds a Bézier curve to a report, and returns a **Shape** object that represents the curve.

## Syntax

_expression_. `AddCurve` _(SafeArrayOfPoints)_

_expression_ A variable that represents a **[Shapes](Project.Shapes.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _SafeArrayOfPoints_|Required|**Variant**|An array of coordinate pairs that specifies the vertices and control points of the curve.|
| _SafeArrayOfPoints_|Required|**Variant**||

## Return value

 **Shape**


## Remarks

For the  _SafeArrayOfPoints_ parameter, the first point you specify is the starting vertex, and the next two points are control points for the first Bézier segment. Then, for each additional segment of the curve, you specify a vertex and two control points. The last point you specify is the ending vertex for the curve. Note that you must always specify 3 _n_ + 1 points, where _n_ is the number of segments in the curve.


## Example

The following example creates a curve that has seven vertices, starting at the upper-left corner of the report. The curve is set to a yellow-green line that is two points wide.


```vb
Sub AddBezierCurve()
    Dim shapeReport As Report
    Dim reportName As String
    Dim curveShape As shape
    
    ' Add a report.
    reportName = "Curve report"
    Set shapeReport = ActiveProject.Reports.Add(reportName)

    Dim pts(1 To 7, 1 To 2) As Single
    pts(1, 1) = 0
    pts(1, 2) = 0
    pts(2, 1) = 72
    pts(2, 2) = 72
    pts(3, 1) = 100
    pts(3, 2) = 40
    pts(4, 1) = 20
    pts(4, 2) = 50
    pts(5, 1) = 90
    pts(5, 2) = 120
    pts(6, 1) = 60
    pts(6, 2) = 30
    pts(7, 1) = 150
    pts(7, 2) = 90

    Set curveShape = shapeReport.Shapes.AddCurve(pts)

    With curveShape
        .Line.Weight = 2
        .Line.ForeColor.RGB = &H1FFAA
    End With
End Sub
```


## See also


[Shapes Object](Project.shapes.md)
[Shape Object](Project.shape.md)
[Line Property](Project.shape.line.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]