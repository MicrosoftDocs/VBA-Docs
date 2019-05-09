---
title: Series.Trendlines method (Excel)
keywords: vbaxl10.chm578107
f1_keywords:
- vbaxl10.chm578107
ms.prod: excel
api_name:
- Excel.Series.Trendlines
ms.assetid: d42609e1-011c-6cb3-286d-192284cd8ab8
ms.date: 06/08/2017
localization_priority: Normal
---


# Series.Trendlines method (Excel)

Returns an object that represents a single trendline (a  **[Trendline](Excel.Trendline(object).md)** object) or a collection of all the trendlines (a **[Trendlines](Excel.Trendlines(object).md)** collection) for the series.


## Syntax

_expression_. `Trendlines` (_Index_)

_expression_ A variable that represents a [Series](Excel.Series-graph-object.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Index_|Optional| **Variant**|The name or number of the trendline.|

## Return value

Object


## Example

This example adds a linear trendline to series one on Chart1.


```vb
Charts("Chart1").SeriesCollection(1).Trendlines.Add Type:=xlLinear
```


## See also


[Series Object](Excel.Series(object).md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]