---
title: ChartArea object (PowerPoint)
keywords: vbapp10.chm687000
f1_keywords:
- vbapp10.chm687000
ms.prod: powerpoint
api_name:
- PowerPoint.ChartArea
ms.assetid: 2c8bd84e-18e7-6417-de4d-d643064e20f5
ms.date: 06/08/2017
localization_priority: Normal
---


# ChartArea object (PowerPoint)

Represents the chart area of a chart. 


## Remarks

The chart area includes everything, including the plot area. However, the  **[PlotArea](PowerPoint.PlotArea.md)** object has its own formatting, so formatting the plot area does not format the chart area.

Use the  **[ChartArea](PowerPoint.Chart.ChartArea.md)** property to return the **ChartArea** object.


## Example




> [!NOTE] 
> Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example turns off the border for the chart area in the first chart of the active document.




```vb
With ActiveDocument.InlineShapes(1).Chart

    ChartArea.Format.Line.Visible = False

End With
```


## See also


[PowerPoint Object Model Reference](overview/PowerPoint/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]