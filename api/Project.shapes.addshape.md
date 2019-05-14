---
title: Shapes.AddShape method (Project)
ms.prod: project-server
ms.assetid: 58af0a51-a455-5c9a-1cae-e56dc67a08a5
ms.date: 06/08/2017
localization_priority: Normal
---


# Shapes.AddShape method (Project)
Adds a shape of the specified AutoShape type to a report, and returns a  **Shape** object that represents the new shape.

## Syntax

_expression_. `AddShape` _(Type,_ _Left,_ _Top,_ _Width,_ _Height)_

_expression_ A variable that represents a **[Shapes](Project.Shapes.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Type_|Required|**MsoAutoShapeType**|Specifies the type of AutoShape to create.|
| _Left_|Required|**Single**|The position, in [points](../language/glossary/vbe-glossary.md#point), of the left edge of the AutoShape.|
| _Top_|Required|**Single**|The position, in [points](../language/glossary/vbe-glossary.md#point), of the top edge of the AutoShape.|
| _Width_|Required|**Single**|The width, in [points](../language/glossary/vbe-glossary.md#point), of the AutoShape.|
| _Height_|Required|**Single**|The height, in [points](../language/glossary/vbe-glossary.md#point), of the AutoShape.|
| _Type_|Required|MSOAUTOSHAPETYPE||
| _Left_|Required|FLOAT||
| _Top_|Required|FLOAT||
| _Width_|Required|FLOAT||
| _Height_|Required|FLOAT||
|Name|Required/Optional|Data type|Description|

## Return value

 **Shape**


## Remarks

To change the type of an AutoShape, set the  **AutoShapeType** property.


## Example

The following example creates a report that contains two cloud shapes, and then changes the second cloud shape to a yellow speech balloon.


```vb
Sub TestShapes()
    Dim shapeReport As Report
    Dim reportName As String
    
    ' Add a report.
    reportName = "Shape report"
    Set shapeReport = ActiveProject.Reports.Add(reportName)

    ' Add two clouds.
    Dim cloudShape1 As shape
    Dim cloudShape2 As shape
    Set cloudShape1 = shapeReport.Shapes.AddShape(msoShapeCloud, 20, 20, 100, 60)
    Set cloudShape2 = shapeReport.Shapes.AddShape(msoShapeCloud, 100, 200, 60, 100)
    
    ' Change the blue cloud to a yellow speech balloon.
    cloudShape2.AutoShapeType = msoShapeBalloon
    cloudShape2.Fill.ForeColor.RGB = &H80FFFF
End Sub
```


## See also


[Shapes Object](Project.shapes.md)
[Shape Object](Project.shape.md)
[AutoShapeType Property](Project.shape.autoshapetype.md)
[MsoAutoShapeType enumeration (Office)](https://msdn.microsoft.com/library/office/ff862770%28v=office.15%29)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]