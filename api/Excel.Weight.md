---
title: Weight property (Excel Graph)
keywords: vbagr10.chm65656
f1_keywords:
- vbagr10.chm65656
ms.prod: excel
api_name:
- Excel.Weight
ms.assetid: 59a3b106-5811-f082-d9cf-c21f2945da31
ms.date: 04/12/2019
localization_priority: Normal
---


# Weight property (Excel Graph)

Returns or sets the weight of the border. Read/write **[XlBorderWeight](excel.xlborderweight.md)**.

## Syntax

_expression_.**Weight**

_expression_ Required. An expression that returns one of the objects in the **Applies To** list.

## Example

This example sets the border weight for the chart area.

```vb
myChart.ChartArea.Border.Weight = xlMedium
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]