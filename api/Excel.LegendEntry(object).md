---
title: LegendEntry Object (Excel)
keywords: vbaxl10.chm585072
f1_keywords:
- vbaxl10.chm585072
ms.prod: excel
api_name:
- Excel.LegendEntry
ms.assetid: ebe8c35c-87b4-11e6-0675-b8bcc8c668a5
ms.date: 06/08/2017
---


# LegendEntry Object (Excel)

Represents a legend entry in a chart legend.


## Remarks

 The **LegendEntry** object is a member of the **[LegendEntries](Excel.LegendEntries(object).md)** collection. The **LegendEntries** collection contains all the **LegendEntry** objects in the legend.

Each legend entry has two parts: the text of the entry, which is the name of the series associated with the legend entry; and an entry marker, which visually links the legend entry with its associated series or trendline in the chart. Formatting properties for the entry marker and its associated series or trendline are contained in the  **[LegendKey](Excel.LegendKey(object).md)** object.

The text of a legend entry cannot be changed.  **LegendEntry** objects support font formatting, and they can be deleted. No pattern formatting is supported for legend entries. The position and size of entries is fixed.

There's no direct way to return the series or trendline corresponding to the legend entry.

After legend entries have been deleted, the only way to restore them is to remove and recreate the legend that contained them by setting the  **[HasLegend](Excel.Chart.HasLegend.md)** property for the chart to **False** and then back to **True** .


## Example

Use  **[LegendEntries](Excel.Legend.LegendEntries.md)** ( _index_ ), where _index_ is the legend entry index number, to return a single **LegendEntry** object. You cannot return legend entries by name.



The index number represents the position of the legend entry in the legend.  `LegendEntries(1)` is at the top of the legend, and `LegendEntries(LegendEntries.Count)` is at the bottom. The following example changes the font for the text of the legend entry at the top of the legend (this is usually the legend for series one) in embedded chart one on the worksheet named "Sheet1."




```vb
Worksheets("sheet1").ChartObjects(1).Chart _ 
 .Legend.LegendEntries(1).Font.Italic = True
```


## See also


[Excel Object Model Reference](overview/Excel/object-model.md)


