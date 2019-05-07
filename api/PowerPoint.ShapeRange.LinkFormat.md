---
title: ShapeRange.LinkFormat property (PowerPoint)
keywords: vbapp10.chm548045
f1_keywords:
- vbapp10.chm548045
ms.prod: powerpoint
api_name:
- PowerPoint.ShapeRange.LinkFormat
ms.assetid: aa2f91d3-b3fd-9834-b189-ec7fc783db6d
ms.date: 06/08/2017
localization_priority: Normal
---


# ShapeRange.LinkFormat property (PowerPoint)

Returns a  **[LinkFormat](PowerPoint.LinkFormat.md)** object that contains the properties that are unique to linked OLE objects. Read-only.


## Syntax

_expression_. `LinkFormat`

_expression_ A variable that represents a **[ShapeRange](PowerPoint.ShapeRange.md)** object.


## Return value

LinkFormat


## Example

This example updates the links between any OLE objects on slide one in the active presentation and their source files.


```vb
For Each sh In ActivePresentation.Slides(1).Shapes

    If sh.Type = msoLinkedOLEObject Then

        With sh.LinkFormat

            .Update

        End With

    End If

Next
```


## See also


[ShapeRange Object](PowerPoint.ShapeRange.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]