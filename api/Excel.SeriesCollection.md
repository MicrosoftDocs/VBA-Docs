---
title: SeriesCollection object (Excel)
keywords: vbaxl10.chm579072
f1_keywords:
- vbaxl10.chm579072
ms.prod: excel
api_name:
- Excel.SeriesCollection
ms.assetid: 93aa1f0b-4939-8c60-a444-2f791e8ce144
ms.date: 04/02/2019
localization_priority: Normal
---


# SeriesCollection object (Excel)

A collection of all the **[Series](Excel.Series(object).md)** objects in the specified chart or chart group.


## Remarks

Use the **[SeriesCollection](Excel.Chart.SeriesCollection.md)** method of the **Chart** property to return the **SeriesCollection** collection.


## Example

The following example adds the data in cells C1:C10 on worksheet one to an existing series in the series collection in embedded chart one.

```vb
Worksheets(1).ChartObjects(1).Chart. _ 
 SeriesCollection.Extend Worksheets(1).Range("c1:c10")
```

<br/>

Use the **Add** method to create a new series and add it to the chart. The following example adds the data from cells A1:A19 as a new series on the chart sheet named Chart1.

```vb
Charts("chart1").SeriesCollection.Add _ 
 source:=Worksheets("sheet1").Range("a1:a19")
```

<br/>

Use **SeriesCollection** (_index_), where _index_ is the series index number or name, to return a single **Series** object. The following example sets the color of the interior for the first series in embedded chart one on Sheet1.

```vb
Worksheets("sheet1").ChartObjects(1).Chart. _ 
 SeriesCollection(1).Interior.Color = RGB(255, 0, 0)
```


## Methods

- [Add](Excel.SeriesCollection.Add.md)
- [Extend](Excel.SeriesCollection.Extend.md)
- [Item](Excel.SeriesCollection.Item.md)
- [NewSeries](Excel.SeriesCollection.NewSeries.md)
- [Paste](Excel.SeriesCollection.Paste.md)

## Properties

- [Application](Excel.SeriesCollection.Application.md)
- [Count](Excel.SeriesCollection.Count.md)
- [Creator](Excel.SeriesCollection.Creator.md)
- [Parent](Excel.SeriesCollection.Parent.md)


## See also

- [Excel Object Model Reference](overview/excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
