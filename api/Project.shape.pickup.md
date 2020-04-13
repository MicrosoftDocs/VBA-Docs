---
title: Shape.PickUp method (Project)
ms.prod: project-server
ms.assetid: 954390b6-8992-f239-d891-467ec732b0b0
ms.date: 06/08/2017
localization_priority: Normal
---


# Shape.PickUp method (Project)
Copies the formatting of a shape.

## Syntax

_expression_.**PickUp**

_expression_ A variable that represents a **[Shape](Project.Shape.md)** object.


## Return value

 **Nothing**


### Remarks

Use the **[Apply](Project.shape.apply.md)** method to apply copied formatting to another shape.


## Example

The following example creates two cylindrical shapes, colors the first shape red, copies the formatting of the first shape, and then applies it to the second shape.


```vb
Sub ApplyShapeFormat()
    Dim theReport As Report
    Dim shp1 As shape
    Dim shp2 As shape
    Dim reportName As String
    Dim sRange As ShapeRange
    
    reportName = "Apply Report"
    
    Set theReport = ActiveProject.Reports.Add(reportName)
    Set shp1 = theReport.Shapes.AddShape(msoShapeCan, 10, 30, 100, 100)
    shp1.Name = "Shape 1"
    shp1.Fill.ForeColor.RGB = &H1010FF  ' Red color.
    
    Set shp2 = theReport.Shapes.AddShape(msoShapeCan, 30, 140, 100, 100)
    shp2.Name = "Shape 2"               ' Blue default color.
    
    With theReport
        .Shapes("Shape 1").PickUp
        .Shapes("Shape 2").Apply
    End With
End Sub
```


## See also


[Shape Object](Project.shape.md)
[Apply Method](Project.shape.apply.md)
[ShapeRange.Pickup Method](Project.shaperange.pickup.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]