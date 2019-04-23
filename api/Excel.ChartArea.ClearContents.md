---
title: ChartArea.ClearContents method (Excel)
keywords: vbaxl10.chm620078
f1_keywords:
- vbaxl10.chm620078
ms.prod: excel
api_name:
- Excel.ChartArea.ClearContents
ms.assetid: 3c3c07a0-9dc1-6019-5262-e1acba7917a1
ms.date: 04/20/2019
localization_priority: Normal
---


# ChartArea.ClearContents method (Excel)

Clears the data from a chart but leaves the formatting.


## Syntax

_expression_.**ClearContents**

_expression_ A variable that represents a **[ChartArea](Excel.ChartArea(object).md)** object.


## Return value

Variant


## Example

This example clears the chart data from Chart1 but leaves the formatting intact.

```vb
Charts("Chart1").ChartArea.ClearContents
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]