---
title: VerticalAlignment property (Excel Graph)
keywords: vbagr10.chm65673
f1_keywords:
- vbagr10.chm65673
ms.prod: excel
ms.assetid: 0021576c-89c5-79ea-bfad-2e29ee9425ae
ms.date: 04/12/2019
localization_priority: Normal
---


# VerticalAlignment property (Excel Graph)

Returns or sets the vertical alignment of the specified object. Required **[XlVAlign](excel.xlvalign.md)**.

## Syntax

_expression_.**VerticalAlignment**

_expression_ Required. An expression that returns one of the objects in the **Applies To** list.

## Remarks

Some of these constants may not be available to you depending on the language support (U.S. English, for example) that you've selected or installed.


## Example

This example centers the chart title vertically.

```vb
myChart.ChartTitle.VerticalAlignment = xlCenter
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]