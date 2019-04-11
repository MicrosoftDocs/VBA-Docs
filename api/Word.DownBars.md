---
title: DownBars object (Word)
ms.prod: word
api_name:
- Word.DownBars
ms.assetid: d0cf170e-0c58-2d01-a4b2-1eaf65dbfa3c
ms.date: 06/08/2017
localization_priority: Normal
---


# DownBars object (Word)

Represents the down bars in a chart group.


## Remarks

 Down bars connect points on the first series in the chart group with lower values on the last series (the lines go down from the first series). Only 2D line groups that contain at least two series can have down bars. This object is not a collection. There is no object that represents a single down bar; you either enable up bars and down bars for all points in a chart group or you disable them.

If the  **[HasUpDownBars](Word.ChartGroup.HasUpDownBars.md)** property is **False**, most properties of the **DownBars** object are disabled.


## Example

Use the  **[DownBars](Word.ChartGroup.DownBars.md)** property to return the **DownBars** object. The following example enables up and down bars for chart group one of the first chart in the active document. The example then sets the up bar color to blue and the down bar color to red.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 With .Chart.ChartGroups(1) 
 .HasUpDownBars = True 
 .UpBars.Interior.Color = RGB(0, 0, 255) 
 .DownBars.Interior.Color = RGB(255, 0, 0) 
 End With 
 End If 
End With 

```


## See also


[Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]