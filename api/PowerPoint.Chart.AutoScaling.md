---
title: Chart.AutoScaling property (PowerPoint)
keywords: vbapp10.chm684015
f1_keywords:
- vbapp10.chm684015
ms.prod: powerpoint
api_name:
- PowerPoint.Chart.AutoScaling
ms.assetid: 330a185a-713a-409a-704e-3b163394aa92
ms.date: 06/08/2017
localization_priority: Normal
---


# Chart.AutoScaling property (PowerPoint)

 **True** if Microsoft Word scales a 3D chart so that it is closer in size to the equivalent 2D chart. The **[RightAngleAxes](PowerPoint.Chart.RightAngleAxes.md)** property must be **True**. Read/write **Boolean**.


## Syntax

_expression_. `AutoScaling`

_expression_ A variable that represents a **[Chart](PowerPoint.Chart.md)** object.


## Example




> [!NOTE] 
> Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example automatically scales the first chart in the active document. The example should be run on a 3D chart.




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        .Chart.RightAngleAxes = True

        .Chart.AutoScaling = True

    End If

End With
```


## See also


[Chart Object](PowerPoint.Chart.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]