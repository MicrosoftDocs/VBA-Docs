---
title: Shape.AutoShapeType property (PowerPoint)
keywords: vbapp10.chm547016
f1_keywords:
- vbapp10.chm547016
ms.prod: powerpoint
api_name:
- PowerPoint.Shape.AutoShapeType
ms.assetid: 99c8e48a-2e0e-0c5a-fb78-91790c668bd7
ms.date: 06/08/2017
localization_priority: Normal
---


# Shape.AutoShapeType property (PowerPoint)

Returns or sets the shape type for the specified  **Shape** object, which must represent an AutoShape other than a line, freeform drawing, or connector. Read/write.


## Syntax

_expression_.**AutoShapeType**

_expression_ A variable that represents a **[Shape](PowerPoint.Shape.md)** object.


## Return value

[MsoAutoShapeType](Office.MsoAutoShapeType.md)


## Remarks

Use the  **Type** property of the **[ConnectorFormat](PowerPoint.ConnectorFormat.md)** object to set or return the connector type.


## Example

This example replaces all 16-point stars with 32-point stars in _myDocument_.


```vb
Set myDocument = ActivePresentation.Slides(1) 
For Each s In myDocument.Shapes 
    If s.AutoShapeType = msoShape16pointStar Then 
        s.AutoShapeType = msoShape32pointStar 
    End If 
Next
```


> [!NOTE] 
> When you change the type of a shape, the shape retains its size, color, and other attributes.


## See also


[Shape Object](PowerPoint.Shape.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]