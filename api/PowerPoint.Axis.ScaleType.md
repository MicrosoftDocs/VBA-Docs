---
title: Axis.ScaleType Property (PowerPoint)
keywords: vbapp10.chm682026
f1_keywords:
- vbapp10.chm682026
ms.prod: powerpoint
api_name:
- PowerPoint.Axis.ScaleType
ms.assetid: baf40097-28a4-c2ec-fea9-2ce971f72ed5
ms.date: 06/08/2017
localization_priority: Normal
---


# Axis.ScaleType Property (PowerPoint)

Returns or sets the value axis scale type. Read/write  **[xlScaleType](PowerPoint.XlScaleType.md)**.


## Syntax

 _expression_. `ScaleType`

_expression_ A variable that represents an '[Axis](PowerPoint.Axis.md)' object.


## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example sets the value axis for the first chart in the active document to use a logarithmic scale.




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        .Chart.Axes(xlValue).ScaleType = xlScaleLogarithmic

    End If

End With


```


## See also


[Axis Object](PowerPoint.Axis.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]