---
title: TickLabels object (PowerPoint)
keywords: vbapp10.chm719000
f1_keywords:
- vbapp10.chm719000
ms.prod: powerpoint
api_name:
- PowerPoint.TickLabels
ms.assetid: 2ba878bf-3a76-1350-2bd4-615c2520f042
ms.date: 06/08/2017
localization_priority: Normal
---


# TickLabels object (PowerPoint)

Represents the tick-mark labels associated with tick marks on a chart axis.


## Remarks

This object is not a collection. There is no object that represents a single tick-mark label; you must return all the tick-mark labels as a unit.

Tick-mark label text for the category axis comes from the name of the associated category in the chart. The default tick-mark label text for the category axis is the number that indicates the position of the category relative to the left end of this axis. To change the number of unlabeled tick marks between tick-mark labels, you must change the  **[TickLabelSpacing](PowerPoint.Axis.TickLabelSpacing.md)** property for the category axis.

Tick-mark label text for the value axis is calculated based on the  **[MajorUnit](PowerPoint.Axis.MajorUnit.md)**, **[MinimumScale](PowerPoint.Axis.MinimumScale.md)**, and **[MaximumScale](PowerPoint.Axis.MaximumScale.md)** properties of the value axis. To change the tick-mark label text for the value axis, you must change the values of these properties.


## Example




> [!NOTE] 
> Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

Use the  **[TickLabels](PowerPoint.Axis.TickLabels.md)** property to return the **TickLabels** object. The following example sets the number format for the tick-mark labels on the value axis for the first chart in the active document.




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        .Chart.Axes(xlValue).TickLabels.NumberFormat = "0.00"

    End If

End With
```


## See also


[PowerPoint Object Model Reference](overview/PowerPoint/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]