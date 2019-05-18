---
title: Trendlines method (Excel Graph)
keywords: vbagr10.chm3077635
f1_keywords:
- vbagr10.chm3077635
ms.prod: excel
api_name:
- Excel.Trendlines
ms.assetid: 2379333d-1cca-bd04-2dec-170bd5d40f67
ms.date: 04/09/2019
localization_priority: Normal
---


# Trendlines method (Excel Graph)

Returns an object that represents a single trendline or a collection of all the trendlines for the series.

## Syntax

_expression_.**Trendlines** (_Index_)

_expression_ Required. An expression that returns a **[Trendlines](Excel.trendlines(collection).md)** collection.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_Index_ | Optional |**Variant**| The name or number of the trendline.|

## Example

This example adds a linear trendline to series one.

```vb
myChart.SeriesCollection(1).Trendlines.Add Type:=xlLinear
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]