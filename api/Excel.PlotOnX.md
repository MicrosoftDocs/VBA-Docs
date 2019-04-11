---
title: PlotOnX property (Excel Graph)
keywords: vbagr10.chm67311
f1_keywords:
- vbagr10.chm67311
ms.prod: excel
api_name:
- Excel.PlotOnX
ms.assetid: 66102cce-e4af-4b0c-d168-ea63f3bc0f30
ms.date: 04/11/2019
localization_priority: Normal
---


# PlotOnX property (Excel Graph)

Returns or sets the index of the data sheet row whose contents are to be used as the X-axis values in the specified X-Y scatter chart. Read/write **Long**.

## Syntax

_expression_.**PlotOnX**

_expression_ Required. An expression that returns one of the objects in the **Applies To** list.

## Example

This example sets row 0 as the row whose contents will be plotted as values on the X-axis in _myChart_.

```vb
myChart.PlotOnX = 0 

```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]