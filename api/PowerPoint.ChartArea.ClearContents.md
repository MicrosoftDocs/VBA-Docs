---
title: ChartArea.ClearContents method (PowerPoint)
keywords: vbapp10.chm65649
f1_keywords:
- vbapp10.chm65649
ms.prod: powerpoint
api_name:
- PowerPoint.ChartArea.ClearContents
ms.assetid: 7cb3e9a9-e808-ed80-c55e-de422d19d9e3
ms.date: 06/08/2017
localization_priority: Normal
---


# ChartArea.ClearContents method (PowerPoint)

Clears the data from a chart but leaves the formatting.


## Syntax

_expression_.**ClearContents**

_expression_ A variable that represents a '[ChartArea](PowerPoint.ChartArea.md)' object.


## Example




> [!NOTE] 
> Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example clears the chart data from the first chart in the active document but leaves the formatting intact.




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        .Chart.ChartArea.ClearContents

    End If

End With
```


## See also


[ChartArea Object](PowerPoint.ChartArea.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]