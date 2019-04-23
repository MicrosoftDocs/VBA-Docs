---
title: Series object (Excel Graph)
keywords: vbagr10.chm131115
f1_keywords:
- vbagr10.chm131115
ms.prod: excel
api_name:
- Excel.Series
ms.assetid: c4446d04-9a3a-4f95-7b3f-adaf1ad2252c
ms.date: 04/06/2019
localization_priority: Normal
---


# Series object (Excel Graph)

Represents a series in the specified chart. The **Series** object is a member of the **[SeriesCollection](Excel.seriescollection(collection).md)** collection.


## Remarks

Use **SeriesCollection** (_index_), where _index_ is the series' index number or name, to return a single **Series** object. 

The series index number indicates the order in which the series are added to the chart. `SeriesCollection(1)` is the first series added to the chart, and `SeriesCollection(SeriesCollection.Count)` is the last one added.


## Example

The following example sets the color of the interior for series one in the chart.

```vb
myChart.SeriesCollection(1).Interior.Color = RGB(255, 0, 0)
```


## See also

- [Excel Graph Visual Basic Reference](overview/excel/graph-visual-basic-reference.md)
- [Excel Object Model Reference](overview/excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]