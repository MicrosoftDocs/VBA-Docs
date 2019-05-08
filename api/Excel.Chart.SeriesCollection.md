---
title: Chart.SeriesCollection method (Excel)
keywords: vbaxl10.chm149141
f1_keywords:
- vbaxl10.chm149141
ms.prod: excel
api_name:
- Excel.Chart.SeriesCollection
ms.assetid: 0a628f00-1ee6-9ff8-dce1-c7aabbdd1a85
ms.date: 04/16/2019
localization_priority: Normal
---


# Chart.SeriesCollection method (Excel)

Returns an object that represents either a single series (a **[Series](Excel.Series(object).md)** object) or a collection of all the series (a **[SeriesCollection](Excel.SeriesCollection.md)** collection) in the chart or chart group.


## Syntax

_expression_.**SeriesCollection** (_Index_)

_expression_ A variable that represents a **[Chart](Excel.Chart(object).md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Index_|Optional| **Variant**|The name or number of the series.|

## Return value

Object


## Example

This example turns on data labels for series one on Chart1.

```vb
Charts("Chart1").SeriesCollection(1).HasDataLabels = True
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
