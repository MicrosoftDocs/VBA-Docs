---
title: Chart.Walls property (PowerPoint)
keywords: vbapp10.chm684047
f1_keywords:
- vbapp10.chm684047
ms.prod: powerpoint
api_name:
- PowerPoint.Chart.Walls
ms.assetid: e4c019c0-41de-988b-b5c7-009fcc0eee15
ms.date: 06/08/2017
localization_priority: Normal
---


# Chart.Walls property (PowerPoint)

Returns the walls of the 3D chart. Read-only  **[Walls](PowerPoint.Walls.md)**.


## Syntax

_expression_.**Walls**

_expression_ A variable that represents a **[Chart](PowerPoint.Chart.md)** object.


## Example




> [!NOTE] 
> Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example sets the color of the wall border of the first chart in the active document to red. You should run the example on a 3D chart.




```vb
With ActiveDocument.InlineShapes(1)
    If .HasChart Then
        .Chart.Walls.Border. _
            ColorIndex = 3
    End If
End With


```


## See also


[Chart Object](PowerPoint.Chart.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]