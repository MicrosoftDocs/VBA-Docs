---
title: MajorTickMark property (Excel Graph)
keywords: vbagr10.chm65562
f1_keywords:
- vbagr10.chm65562
ms.prod: excel
api_name:
- Excel.MajorTickMark
ms.assetid: 26dfa842-1c7d-c2b6-b647-7c110b1d5626
ms.date: 04/11/2019
localization_priority: Normal
---


# MajorTickMark property (Excel Graph)

Returns or sets the type of major tick mark for the specified axis. Read/write **[XlTickMark](excel.xltickmark.md)**.

## Syntax

_expression_.**MajorTickMark**

_expression_ Required. An expression that returns one of the objects in the **Applies To** list.

## Example

This example sets the major tick marks for the value axis to be outside the axis.

```vb
myChart.Axes(xlValue).MajorTickMark = xlTickMarkOutside
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]