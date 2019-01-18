---
title: Series.HasErrorBars property (Excel)
keywords: vbaxl10.chm578089
f1_keywords:
- vbaxl10.chm578089
ms.prod: excel
api_name:
- Excel.Series.HasErrorBars
ms.assetid: 03d9a6e6-8c15-2bdb-1bca-ed9fb95e9cb3
ms.date: 06/08/2017
localization_priority: Normal
---


# Series.HasErrorBars property (Excel)

 **True** if the series has error bars. This property isn't available for 3-D charts. Read/write **Boolean**.


## Syntax

_expression_. `HasErrorBars`

_expression_ A variable that represents a [Series](Excel.Series-graph-object.md) object.


## Example

This example removes error bars from series one in Chart1. The example should be run on a 2-D line chart that has error bars for series one.


```vb
Charts("Chart1").SeriesCollection(1).HasErrorBars = False
```


## See also


[Series Object](Excel.Series(object).md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]