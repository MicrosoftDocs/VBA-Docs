---
title: HorizontalAlignment property (Excel Graph)
keywords: vbagr10.chm65672
f1_keywords:
- vbagr10.chm65672
ms.prod: excel
ms.assetid: 7af45990-24df-8dbf-92ec-a06b9f718f9e
ms.date: 04/11/2019
localization_priority: Normal
---


# HorizontalAlignment property (Excel Graph)

Returns or sets the horizontal alignment for the specified object. Read/write **[XlHAlign](excel.xlhalign.md)**.

## Syntax

_expression_.**HorizontalAlignment**

_expression_ Required. An expression that returns one of the objects in the **Applies To** list.

## Remarks

Some of these constants may not be available to you, depending on the language support (U.S. English, for example) that you've selected or installed.


## Example

This example centers the chart title.

```vb
myChart.ChartTitle.HorizontalAlignment = xlCenter
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]