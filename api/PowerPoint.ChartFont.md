---
title: ChartFont object (PowerPoint)
keywords: vbapp10.chm704000
f1_keywords:
- vbapp10.chm704000
ms.prod: powerpoint
api_name:
- PowerPoint.ChartFont
ms.assetid: 185dfaa0-4ed9-01d2-6584-b0838b50ef8c
ms.date: 06/08/2017
localization_priority: Normal
---


# ChartFont object (PowerPoint)

Contains the font attributes (font name, font size, color, and so on) for an object chart.


## Remarks

If you do not want to format all the text in an  **[AxisTitle](PowerPoint.AxisTitle.md)**, **[ChartTitle](PowerPoint.ChartTitle.md)**, **[DataLabel](PowerPoint.DataLabel.md)**, or **[DisplayUnitLabel](PowerPoint.DisplayUnitLabel.md)** object the same way, use the **Characters** property of that object to first return a subset of the text as a **[ChartCharacters](PowerPoint.ChartCharacters.md)** object. Then use the **[Font](PowerPoint.ChartCharacters.Font.md)** property of the **ChartCharacters** object to return a **ChartFont** object you can use to format the subset of text, as needed.


## Example




> [!NOTE] 
> Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example formats the title of the first chart as bold. Use the  **Font** property to return the **ChartFont** object.




```vb
With ActiveDocument.InlineShapes(1).Chart

    .AxisTitle.Font.Bold = True

End With
```


## See also


[PowerPoint Object Model Reference](overview/PowerPoint/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]