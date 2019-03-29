---
title: FullSeriesCollection object (Excel)
keywords: vbaxl10.chm943072
f1_keywords:
- vbaxl10.chm943072
ms.prod: excel
ms.assetid: 5d7b7e7c-0a74-307b-84f9-56143ceba464
ms.date: 03/30/2019
localization_priority: Normal
---


# FullSeriesCollection object (Excel)

Represents the full set of **[Series](Excel.Series(object).md)** objects in a chart.


## Remarks

The **FullSeriesCollection** object enables you to get a filtered out **Series** object and filter it back in. It also enables you to iterate over the full set of **Series** objects, filtered out or visible, programmatically. 

By having the existing **[SeriesCollection](Excel.SeriesCollection.md)** object contain only the visible series, you can programmatically perform operations on only the visible series. It also prevents Microsoft Excel from breaking existing chart solutions on charts with filtered out data.


## Example

The following example displays a message box with the name of the second **Series** object in the second chart.

```vb
MsgBox Chart(1).FullSeriesCollection.Item(2).Name
```


## Methods

- [Item](Excel.fullseriescollection.item.md)

## Properties

- [Application](Excel.fullseriescollection.application.md)
- [Count](Excel.fullseriescollection.count.md)
- [Creator](Excel.fullseriescollection.creator.md)
- [Parent](Excel.fullseriescollection.parent.md)

## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
