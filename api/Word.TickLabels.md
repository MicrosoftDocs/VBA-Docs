---
title: TickLabels object (Word)
keywords: vbawd10.chm2549
f1_keywords:
- vbawd10.chm2549
ms.prod: word
api_name:
- Word.TickLabels
ms.assetid: d94e90dc-0b0e-f4af-078e-6f2b97729db5
ms.date: 06/08/2017
localization_priority: Normal
---


# TickLabels object (Word)

Represents the tick-mark labels associated with tick marks on a chart axis.


## Remarks

This object is not a collection. There is no object that represents a single tick-mark label; you must return all the tick-mark labels as a unit.

Tick-mark label text for the category axis comes from the name of the associated category in the chart. The default tick-mark label text for the category axis is the number that indicates the position of the category relative to the left end of this axis. To change the number of unlabeled tick marks between tick-mark labels, you must change the  **[TickLabelSpacing](Word.Axis.TickLabelSpacing.md)** property for the category axis.

Tick-mark label text for the value axis is calculated based on the  **[MajorUnit](Word.Axis.MajorUnit.md)**, **[MinimumScale](Word.Axis.MinimumScale.md)**, and **[MaximumScale](Word.Axis.MaximumScale.md)** properties of the value axis. To change the tick-mark label text for the value axis, you must change the values of these properties.


## Example

Use the  **[TickLabels](Word.Axis.TickLabels.md)** property to return the **TickLabels** object. The following example sets the number format for the tick-mark labels on the value axis for the first chart in the active document.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 .Chart.Axes(xlValue).TickLabels.NumberFormat = "0.00" 
 End If 
End With
```


## See also


[Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]