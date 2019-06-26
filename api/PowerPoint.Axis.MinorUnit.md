---
title: Axis.MinorUnit property (PowerPoint)
keywords: vbapp10.chm682023
f1_keywords:
- vbapp10.chm682023
ms.prod: powerpoint
api_name:
- PowerPoint.Axis.MinorUnit
ms.assetid: ff4b4a7b-25b3-974c-dee1-b81f0ec15634
ms.date: 06/08/2017
localization_priority: Normal
---


# Axis.MinorUnit property (PowerPoint)

Returns or sets the minor units on the value axis. Read/write  **Double**.


## Syntax

_expression_. `MinorUnit`

_expression_ A variable that represents an '[Axis](PowerPoint.Axis.md)' object.


## Remarks

Setting this property sets the  **[MinorUnitIsAuto](PowerPoint.Axis.MinorUnitIsAuto.md)** property to **False**.

Use the  **[TickMarkSpacing](PowerPoint.Axis.TickLabelSpacing.md)** property to set tick-mark spacing on the category axis.


## Example




> [!NOTE] 
> Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example sets the major and minor units for the value axis of the first chart in the active document.




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        With .Chart.Axes(xlValue)

            .MajorUnit = 100

            .MinorUnit = 20

        End With

    End If

End With
```


## See also


[Axis Object](PowerPoint.Axis.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]