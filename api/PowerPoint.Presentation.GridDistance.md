---
title: Presentation.GridDistance property (PowerPoint)
keywords: vbapp10.chm583062
f1_keywords:
- vbapp10.chm583062
ms.prod: powerpoint
api_name:
- PowerPoint.Presentation.GridDistance
ms.assetid: 5c4accfe-2467-3d0e-f7f8-3e3c16d8d0ce
ms.date: 06/08/2017
localization_priority: Normal
---


# Presentation.GridDistance property (PowerPoint)

Sets or returns a  **Single** that represents the distance between gridlines. Read/write.


## Syntax

_expression_. `GridDistance`

_expression_ A variable that represents a [Presentation](PowerPoint.Presentation.md) object.


## Return value

Single


## Example

This example displays the gridlines, and then specifies the distance between gridlines and enables the snap to grid setting.


```vb
Sub SetGridLines()

    Application.DisplayGridLines = msoTrue

    With ActivePresentation

        .GridDistance = 18

        .SnapToGrid = msoTrue

    End With

End Sub
```


## See also


[Presentation Object](PowerPoint.Presentation.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]