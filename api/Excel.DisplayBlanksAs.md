---
title: DisplayBlanksAs property (Excel Graph)
keywords: vbagr10.chm3077021
f1_keywords:
- vbagr10.chm3077021
ms.prod: excel
api_name:
- Excel.DisplayBlanksAs
ms.assetid: c2669ad5-9532-ea7c-120c-bc8a15878864
ms.date: 04/10/2019
localization_priority: Normal
---


# DisplayBlanksAs property (Excel Graph)

Returns or sets the way that blank cells are plotted on a chart. Read/write **[XlDisplayBlanksAs](excel.xldisplayblanksas.md)**.

## Syntax

_expression_.**DisplayBlanksAs**

_expression_ Required. An expression that returns one of the objects in the **Applies To** list.

## Example

This example sets Graph to not plot blank cells.

```vb
myChart.DisplayBlanksAs = xlNotPlotted
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]