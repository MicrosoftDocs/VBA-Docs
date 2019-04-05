---
title: LegendEntry object (Excel Graph)
keywords: vbagr10.chm131194
f1_keywords:
- vbagr10.chm131194
ms.prod: excel
api_name:
- Excel.LegendEntry
ms.assetid: a242fdab-ebb4-f5de-04ae-d6b70cea1640
ms.date: 04/06/2019
localization_priority: Normal
---


# LegendEntry object (Excel Graph)

Represents a legend entry in the specified chart legend. The **LegendEntry** object is a member of the **[LegendEntries](Excel.legendentries(collection).md)** collection, which contains all the **LegendEntry** objects in the legend.

## Remarks

Each legend entry has two parts: the text of the entry, which is the name of the series associated with the entry; and an entry marker, which visually links the legend entry with its associated series or trendline in the chart. Formatting properties for the entry marker and its associated series or trendline are contained in the **[LegendKey](Excel.LegendKey-graph-object.md)** object.

You cannot change the text of a legend entry. **LegendEntry** objects support font formatting, and they can be deleted. No pattern formatting is supported for legend entries. The position and size of entries is fixed.

Use **LegendEntries** (_index_), where _index_ is the legend entry's index number, to return a single **LegendEntry** object. You cannot return legend entries by name.

The index number represents the position of the legend entry in the legend. `LegendEntries(1)` is at the top of the legend, and `LegendEntries(LegendEntries.Count)` is at the bottom. 

There's no direct way to return the series or trendline that corresponds to a particular legend entry.

After legend entries have been deleted, the only way to restore them is to remove and then recreate the legend that contained them by setting the **[HasLegend](Excel.HasLegend.md)** property for the chart to **False** and then back to **True**.


## Example

The following example changes the font style for the text of the legend entry at the top of the legend (this is usually the legend for series one).

```vb
myChart.Legend.LegendEntries(1).Font.Italic = True
```


## See also

- [Excel Graph Visual Basic Reference](overview/excel/graph-visual-basic-reference.md)
- [Excel Object Model Reference](overview/excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]