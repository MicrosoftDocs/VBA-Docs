---
title: ChartArea.Clear method (PowerPoint)
ms.prod: powerpoint
api_name:
- PowerPoint.ChartArea.Clear
ms.assetid: fa22b630-405c-f771-faaa-14bdf8d9fa8b
ms.date: 06/08/2017
localization_priority: Normal
---


# ChartArea.Clear method (PowerPoint)

Clears the entire object.


## Syntax

_expression_.**Clear**

_expression_ A variable that represents a '[ChartArea](PowerPoint.ChartArea.md)' object.


## Example




> [!NOTE] 
> Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example clears the chart area (the chart data and formatting) of the first chart in the active document.




```vb
With ActiveDocument.InlineGroups(1)

    If .HasChart Then

        .Chart.ChartArea.Clear

    End If

End With
```


## See also


[ChartArea Object](PowerPoint.ChartArea.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]