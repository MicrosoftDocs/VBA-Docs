---
title: OutlineFont property (Excel Graph)
keywords: vbagr10.chm65757
f1_keywords:
- vbagr10.chm65757
api_name:
- Excel.OutlineFont
ms.assetid: 41075763-8ee7-e6ba-c9a2-7bc718b5405e
ms.date: 04/11/2019
ms.localizationpriority: medium
---


# OutlineFont property (Excel Graph)

**True** if the font is an outline font. Read/write **Variant**.

## Syntax

_expression_.**OutlineFont**

_expression_ Required. An expression that returns one of the objects in the **Applies To** list.


## Remarks

This property has no effect in Windows, but its value is retained (it can be set and returned).


## Example

This example sets the font for the chart title to an outline font.

```vb
myChart.ChartTitle.Font.OutlineFont = True
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]