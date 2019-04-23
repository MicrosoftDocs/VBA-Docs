---
title: Axis.HasDisplayUnitLabel property (PowerPoint)
keywords: vbapp10.chm682044
f1_keywords:
- vbapp10.chm682044
ms.prod: powerpoint
api_name:
- PowerPoint.Axis.HasDisplayUnitLabel
ms.assetid: adbbbb89-55af-12f5-ec67-1e88424f3d81
ms.date: 06/08/2017
localization_priority: Normal
---


# Axis.HasDisplayUnitLabel property (PowerPoint)

 **True** if the label specified by the **[DisplayUnit](PowerPoint.Axis.DisplayUnit.md)** or **[DisplayUnitCustom](PowerPoint.Axis.DisplayUnitCustom.md)** property is displayed on the specified axis. The default is **True**. Read/write **Boolean**.


## Syntax

_expression_.**HasDisplayUnitLabel**

_expression_ A variable that represents an '[Axis](PowerPoint.Axis.md)' object.


## Example

Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example sets the units on the value axis of the first chart in the active document to increments of 500 but keeps the unit label hidden.




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        With .Chart.Axes(xlValue)

            .DisplayUnit = xlCustom

            .DisplayUnitCustom = 500

            .AxisTitle.Caption = "Rebate Amounts"

            .HasDisplayUnitLabel = False

        End With

    End If

End With


```


## See also


[Axis Object](PowerPoint.Axis.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]