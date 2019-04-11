---
title: NameIsAuto property (Excel Graph)
keywords: vbagr10.chm65724
f1_keywords:
- vbagr10.chm65724
ms.prod: excel
api_name:
- Excel.NameIsAuto
ms.assetid: 92a06cde-f3fc-cc5b-9af9-0ec9545b90a8
ms.date: 04/11/2019
localization_priority: Normal
---


# NameIsAuto property (Excel Graph)

**True** if Graph automatically determines the name of the trendline. Read/write **Boolean**.

## Syntax

_expression_.**NameIsAuto**

_expression_ Required. An expression that returns one of the objects in the **Applies To** list.

## Example

This example sets Graph to automatically determine the name for trendline one. The example should be run on a 2D column chart that contains a single series with a trendline.

```vb
myChart.SeriesCollection(1).Trendlines(1).NameIsAuto = True
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]