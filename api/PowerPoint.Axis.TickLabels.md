---
title: Axis.TickLabels property (PowerPoint)
keywords: vbapp10.chm682029
f1_keywords:
- vbapp10.chm682029
ms.prod: powerpoint
api_name:
- PowerPoint.Axis.TickLabels
ms.assetid: 80e39b06-b01d-f817-5357-e6abbbc28e1c
ms.date: 06/08/2017
localization_priority: Normal
---


# Axis.TickLabels property (PowerPoint)

Returns the tick-mark labels for the specified axis. Read-only  **[TickLabels](PowerPoint.TickLabels.md)**.


## Syntax

_expression_. `TickLabels`

_expression_ A variable that represents an '[Axis](PowerPoint.Axis.md)' object.


## Example




> [!NOTE] 
> Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example sets the color of the tick-mark label font for the value axis of the first chart in the active document.




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        .Chart.Axes(xlValue).TickLabels.Font.ColorIndex = 3

    End If

End With
```


## See also


[Axis Object](PowerPoint.Axis.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]