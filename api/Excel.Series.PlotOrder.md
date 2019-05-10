---
title: Series.PlotOrder property (Excel)
keywords: vbaxl10.chm578103
f1_keywords:
- vbaxl10.chm578103
ms.prod: excel
api_name:
- Excel.Series.PlotOrder
ms.assetid: c74ba422-ca4d-db60-02aa-7b512bdd0241
ms.date: 05/11/2019
localization_priority: Normal
---


# Series.PlotOrder property (Excel)

Returns or sets the plot order for the selected series within the chart group. Read/write **Long**.


## Syntax

_expression_.**PlotOrder**

_expression_ A variable that represents a **[Series](Excel.Series(object).md)** object.


## Remarks

You can set plot order only within a chart group (you cannot set the plot order for the entire chart if you have more than one chart type). A chart group is a collection of series with the same chart type.

Changing the plot order of one series will cause the plot orders of the other series in the chart group to be adjusted, as necessary.


## Example

This example makes series two on Chart1 appear third in the plot order. The example should be run on a 2D column chart that contains three or more series.

```vb
Charts("Chart1").ChartGroups(1).SeriesCollection(2).PlotOrder = 3
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]