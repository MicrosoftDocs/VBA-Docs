---
title: DisplayUnitLabel object (PowerPoint)
keywords: vbapp10.chm699000
f1_keywords:
- vbapp10.chm699000
ms.prod: powerpoint
api_name:
- PowerPoint.DisplayUnitLabel
ms.assetid: 4dd4df7d-91c1-9136-2d5b-cdb0794a7716
ms.date: 06/08/2017
localization_priority: Normal
---


# DisplayUnitLabel object (PowerPoint)

Represents a unit label on an axis in the specified chart.


## Remarks

 Unit labels are useful for charting large values (for example, in the millions or billions). You can make the chart more readable by using a single unit label instead of large numbers at each tick mark.


## Example




> [!NOTE] 
> Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

Use the  **[DisplayUnitLabel](PowerPoint.Axis.DisplayUnitLabel.md)** property to return the **DisplayUnitLabel** object. The following example sets the display label caption to "Millions" on the value axis of the first chart in the active document, and then the example turns off automatic font scaling.




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        With .Chart.Axes(xlValue)

            .DisplayUnit = xlMillions

            .HasDisplayUnitLabel = True

            With .DisplayUnitLabel

                .Caption = "Millions"

                .AutoScaleFont = False

            End With

        End With

    End If

End With
```


## See also


[PowerPoint Object Model Reference](overview/PowerPoint/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]