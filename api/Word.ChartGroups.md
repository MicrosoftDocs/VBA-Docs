---
title: ChartGroups object (Word)
ms.prod: word
api_name:
- Word.ChartGroups
ms.assetid: 37136fbd-8740-c817-9666-993bc5d4c847
ms.date: 06/08/2017
localization_priority: Normal
---


# ChartGroups object (Word)

Represents one or more series plotted in a chart with the same format.


## Remarks

 A **ChartGroups** collection is a collection of all the **[ChartGroup](Word.ChartGroup.md)** objects in the specified chart. A chart contains one or more chart groups, each chart group contains one or more series, and each series contains one or more points. For example, a single chart might contain both a line chart group, containing all the series plotted with the line chart format, and a bar chart group, containing all the series plotted with the bar chart format.

 The following example displays the number of chart groups on the first chart of the active document. Use the **[ChartGroups](Word.Chart.ChartGroups.md)** property to return the **ChartGroups** collection.




```vb
MsgBox ActiveDocument.InlineShapes(1).Chart._ 
 ChartGroups.Count
```

The following example adds drop lines to chart group 1 on chart sheet 1. Use  **ChartGroups** (_index_), where _index_ is the chart group index number, to return a single **ChartGroup** object.




```vb
ActiveDocument.InlineShapes(1).Chart._ 
 ChartGroups(1).HasDropLines = True
```


## See also



[Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]