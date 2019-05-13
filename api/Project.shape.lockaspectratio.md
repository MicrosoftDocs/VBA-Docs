---
title: Shape.LockAspectRatio property (Project)
ms.prod: project-server
ms.assetid: b465aad3-d253-d6ce-0f6e-bb0638733647
ms.date: 06/08/2017
localization_priority: Normal
---


# Shape.LockAspectRatio property (Project)
Gets or sets a value that indicates whether the shape retains its original proportions when you resize it; that is, whether the aspect ratio of the shape is locked. Read/write  **[MsoTriState](https://msdn.microsoft.com/library/office/ff860737%28v=office.15%29)**

## Syntax

_expression_.**LockAspectRatio**

_expression_ A variable that represents a **[Shape](Project.Shape.md)** object.


## Remarks

The  **LockAspectRatio** value is **msoTrue** if the shape retains its original proportions when it is resized. If you can change the height and width of the shape independently, the value is **msoFalse**.


## Example

The following example creates two triangles of the same size. The left triangle has the aspect ratio unlocked, and the right triangle has the aspect ratio locked. Figure 1 shows the result when each triangle is resized by the same amount.


```vb
Sub ResizeTriangles()
    Dim shapeReport As Report
    Dim reportName As String
    Dim triangle1 As shape
    Dim triangle2 As shape

    reportName = "Triangle resize report"
    Set shapeReport = ActiveProject.Reports.Add(reportName)
    
    With shapeReport.Shapes
        Set triangle1 = .AddShape(msoShapeIsoscelesTriangle, 10, 10, 100, 100)
        Set triangle2 = .AddShape(msoShapeIsoscelesTriangle, 150, 10, 100, 100)
    End With
    
    triangle1.Select
    triangle1.LockAspectRatio = msoFalse
    triangle1.height = 200
    
    triangle2.Select
    triangle2.LockAspectRatio = msoTrue
    triangle2.height = 200
End Sub
```

In Figure 1, the right shape with the locked aspect ratio is selected.


**Figure 1. Resizing a shape when the aspect ratio is unlocked or locked**

![Resizing a shape with the aspect ratio unlocked](../images/pj15_VBA_LockAspectRatio.gif)


## Property value

 **MSOTRISTATE**


## See also


[Shape Object](Project.shape.md)
[ShapeRange.LockAspectRatio Property](Project.shaperange.lockaspectratio.md)
[MsoTriState](https://msdn.microsoft.com/library/office/ff860737%28v=office.15%29)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]