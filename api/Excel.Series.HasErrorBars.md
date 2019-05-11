---
title: Series.HasErrorBars property (Excel)
keywords: vbaxl10.chm578089
f1_keywords:
- vbaxl10.chm578089
ms.prod: excel
api_name:
- Excel.Series.HasErrorBars
ms.assetid: 03d9a6e6-8c15-2bdb-1bca-ed9fb95e9cb3
ms.date: 05/11/2019
localization_priority: Normal
---


# Series.HasErrorBars property (Excel)

**True** if the series has error bars. This property isn't available for 3D charts. Read/write **Boolean**.


## Syntax

_expression_.**HasErrorBars**

_expression_ A variable that represents a **[Series](Excel.Series(object).md)** object.


## Example

This example removes error bars from series one on Chart1. The example should be run on a 2D line chart that has error bars for series one.

```vb
Charts("Chart1").SeriesCollection(1).HasErrorBars = False
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]