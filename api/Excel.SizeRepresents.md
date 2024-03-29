---
title: SizeRepresents property (Excel Graph)
keywords: vbagr10.chm67188
f1_keywords:
- vbagr10.chm67188
api_name:
- Excel.SizeRepresents
ms.assetid: 54f87d5a-e388-e1d1-8a20-bec820f3449c
ms.date: 04/12/2019
ms.localizationpriority: medium
---


# SizeRepresents property (Excel Graph)

Returns or sets what the bubble size represents on a bubble chart. Read/write **[XlSizeRepresents](excel.xlsizerepresents.md)**.

## Syntax

_expression_.**SizeRepresents**

_expression_ Required. An expression that returns one of the objects in the **Applies To** list.

## Example

This example sets what the bubble size represents for the chart. The example assumes that the chart is a bubble chart.

```vb
myChart.ChartGroups(1).SizeRepresents = xlSizeIsWidth
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]