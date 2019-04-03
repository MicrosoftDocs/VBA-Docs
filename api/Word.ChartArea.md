---
title: ChartArea object (Word)
ms.prod: word
api_name:
- Word.ChartArea
ms.assetid: 7b3384df-f331-033d-4dfa-ee2ff26111c6
ms.date: 06/08/2017
localization_priority: Normal
---


# ChartArea object (Word)

Represents the chart area of a chart. 


## Remarks

The chart area includes everything, including the plot area. However, the  **[PlotArea](Word.PlotArea.md)** object has its own formatting, so formatting the plot area does not format the chart area.

Use the  **[ChartArea](Word.Chart.ChartArea.md)** property to return the **ChartArea** object.


## Example

The following example turns off the border for the chart area in the first chart of the active document.


```vb
With ActiveDocument.InlineShapes(1).Chart 
 ChartArea.Format.Line.Visible = False 
End With
```


## See also



[Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]