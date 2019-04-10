---
title: Model3DFormat.Application property (PowerPoint)
ms.prod: powerpoint
api_name:
- PowerPoint.Model3DFormat.Application
ms.date: 04/11/2019
localization_priority: Normal
---


# Model3DFormat.Application property (PowerPoint)

Returns an **[Application](PowerPoint.Application.md)** object that represents the creator of the specified object.


## Syntax

_expression_.**Application**

_expression_ A variable that represents a **[Model3DFormat](PowerPoint.Model3DFormat.md)** object.


## Return value

Object


## Example

This example displays the name of the application that created each 3D model object on slide one in the active presentation.

```vb
For Each shp3DModel In ActivePresentation.Slides(1).Shapes

    If shp3DModel.Type = mso3DModel Then

        MsgBox shpOle.Model3D.Application.Name

    End If

Next
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]