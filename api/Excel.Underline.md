---
title: Underline property (Excel Graph)
keywords: vbagr10.chm65642
f1_keywords:
- vbagr10.chm65642
ms.prod: excel
api_name:
- Excel.Underline
ms.assetid: 82eb4816-bf37-8a6c-046c-a38ea5c275c2
ms.date: 04/12/2019
localization_priority: Normal
---


# Underline property (Excel Graph)

Returns or sets the type of underline applied to the font. Required **[XlUnderlineStyle](excel.xlunderlinestyle.md)**.

## Syntax

_expression_.**Underline**

_expression_ Required. An expression that returns one of the objects in the **Applies To** list.

## Example

This example sets the font in the chart title to a single underline.

```vb
myChart.ChartTitle.Font.Underline = xlUnderlineStyleSingle
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]