---
title: LegendEntries object (Excel)
keywords: vbaxl10.chm587072
f1_keywords:
- vbaxl10.chm587072
ms.prod: excel
api_name:
- Excel.LegendEntries
ms.assetid: 51d98149-b90b-432b-7771-0815a0e89655
ms.date: 03/30/2019
localization_priority: Normal
---


# LegendEntries object (Excel)

A collection of all the **[LegendEntry](Excel.LegendEntry(object).md)** objects in the specified chart legend.


## Remarks

Each legend entry has two parts: the text of the entry, which is the name of the series or trendline associated with the legend entry; and the entry marker, which visually links the legend entry with its associated series or trendline in the chart. The formatting properties for the entry marker and its associated series or trendline are contained in the **[LegendKey](Excel.LegendKey(object).md)** object.


## Example

Use the **[LegendEntries](Excel.Legend.LegendEntries.md)** method of the **Legend** object to return the **LegendEntries** collection. 

The following example loops through the collection of legend entries in embedded chart one and changes their font color.

```vb
With Worksheets("sheet1").ChartObjects(1).Chart.Legend 
 For i = 1 To .LegendEntries.Count 
 .LegendEntries(i).Font.ColorIndex = 5 
 Next 
End With
```

<br/>

Use **LegendEntries** (_index_), where _index_ is the legend entry index number, to return a single **LegendEntry** object. You cannot return legend entries by name.

The index number represents the position of the legend entry in the legend. `LegendEntries(1)` is at the top of the legend; `LegendEntries(LegendEntries.Count)` is at the bottom. 

The following example changes the font style for the text of the legend entry at the top of the legend (this is usually the legend for series one) in embedded chart one to italic.

```vb
Worksheets("sheet1").ChartObjects(1).Chart _ 
 .Legend.LegendEntries(1).Font.Italic = True
```


## Methods

- [Item](Excel.LegendEntries.Item.md)

## Properties

- [Application](Excel.LegendEntries.Application.md)
- [Count](Excel.LegendEntries.Count.md)
- [Creator](Excel.LegendEntries.Creator.md)
- [Parent](Excel.LegendEntries.Parent.md)


## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
