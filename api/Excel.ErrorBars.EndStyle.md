---
title: ErrorBars.EndStyle property (Excel)
keywords: vbaxl10.chm624079
f1_keywords:
- vbaxl10.chm624079
ms.prod: excel
api_name:
- Excel.ErrorBars.EndStyle
ms.assetid: 865c1da8-1231-5290-c737-c0415615a0ea
ms.date: 06/08/2017
localization_priority: Normal
---


# ErrorBars.EndStyle property (Excel)

Returns or sets the end style for the error bars. Can be one of the following  **[xlEndStyleCap](Excel.XlEndStyleCap.md)** constants: **xlCap** or **xlNoCap**. Read/write **Long**.


## Syntax

_expression_. `EndStyle`

_expression_ A variable that represents an [ErrorBars](Excel.ErrorBars-graph-property.md) object.


## Example

This example sets the end style for the error bars for series one in Chart1. The example should be run on a 2-D line chart that has Y error bars for the first series.


```vb
Charts("Chart1").SeriesCollection(1).ErrorBars.EndStyle = xlCap
```


## See also


[ErrorBars Object](Excel.ErrorBars(object).md)

