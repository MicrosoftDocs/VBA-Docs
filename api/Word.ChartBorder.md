---
title: ChartBorder object (Word)
keywords: vbawd10.chm931
f1_keywords:
- vbawd10.chm931
ms.prod: word
api_name:
- Word.ChartBorder
ms.assetid: eea90670-c599-2ec8-5b7b-c946a4bcd638
ms.date: 06/08/2017
localization_priority: Normal
---


# ChartBorder object (Word)

Represents the border of an object.


## Remarks

Most bordered objects have a border that is treated as a single entity, regardless of how many sides it has. The entire border must be returned as a unit. To return a  **Border** object, use the **Border** property for the particular bordered object (for example, the **[Border](Word.Trendline.Border.md)** property of a **[TrendLine](Word.Trendline.md)** object).


## Example

 The following example changes the type and line style of a trendline on the active chart.


```vb
With ActiveDocument.InlineShapes(1).Chart.SeriesCollection(1).Trendlines(1) 
 .Type = xlLinear 
 .Border.LineStyle = xlDash 
End With
```


## See also


[Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]