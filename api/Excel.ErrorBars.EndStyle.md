---
title: ErrorBars.EndStyle property (Excel)
keywords: vbaxl10.chm624079
f1_keywords:
- vbaxl10.chm624079
ms.prod: excel
api_name:
- Excel.ErrorBars.EndStyle
ms.assetid: 865c1da8-1231-5290-c737-c0415615a0ea
ms.date: 04/26/2019
localization_priority: Normal
---


# ErrorBars.EndStyle property (Excel)

Returns or sets the end style for the error bars. Can be one of the following **[XlEndStyleCap](Excel.XlEndStyleCap.md)** constants: **xlCap** or **xlNoCap**. Read/write **Long**.


## Syntax

_expression_.**EndStyle**

_expression_ A variable that represents an **[ErrorBars](excel.errorbars(object).md)** object.


## Example

This example sets the end style for the error bars for series one on Chart1. The example should be run on a 2D line chart that has Y error bars for the first series.

```vb
Charts("Chart1").SeriesCollection(1).ErrorBars.EndStyle = xlCap
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]