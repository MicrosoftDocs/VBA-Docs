---
title: LegendKey object (Excel)
keywords: vbaxl10.chm589072
f1_keywords:
- vbaxl10.chm589072
ms.prod: excel
api_name:
- Excel.LegendKey
ms.assetid: 2d806a8f-2fed-e6f6-bb76-7339fa692cbb
ms.date: 06/08/2017
localization_priority: Normal
---


# LegendKey object (Excel)

Represents a legend key in a chart legend.


## Remarks

 Each legend key is a graphic that visually links a legend entry with its associated series or trendline in the chart. The legend key is linked to its associated series or trendline in such a way that changing the formatting of one simultaneously changes the formatting of the other.


## Example

Use the  **[LegendKey](Excel.LegendEntry.LegendKey.md)** property to return the **LegendKey** object. The following example changes the marker background color for the legend entry at the top of the legend for embedded chart one on the worksheet named "Sheet1." This simultaneously changes the format of every point in the series associated with this legend entry. The associated series must support data markers.


```vb
Worksheets("sheet1").ChartObjects(1).Chart _ 
 .Legend.LegendEntries(1).LegendKey.MarkerBackgroundColorIndex = 5
```


## See also


[Excel Object Model Reference](overview/Excel/object-model.md)


