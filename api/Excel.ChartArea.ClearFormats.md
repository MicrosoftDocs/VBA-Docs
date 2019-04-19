---
title: ChartArea.ClearFormats method (Excel)
keywords: vbaxl10.chm620082
f1_keywords:
- vbaxl10.chm620082
ms.prod: excel
api_name:
- Excel.ChartArea.ClearFormats
ms.assetid: 0af0bba7-6fb8-d221-7b1f-ba7c40ae1687
ms.date: 04/20/2019
localization_priority: Normal
---


# ChartArea.ClearFormats method (Excel)

Clears the formatting of the object.


## Syntax

_expression_.**ClearFormats**

_expression_ A variable that represents a **[ChartArea](Excel.ChartArea(object).md)** object.


## Return value

Variant


## Example

This example clears the formatting from embedded chart one on Sheet1.

```vb
Worksheets("Sheet1").ChartObjects(1).Chart.ChartArea.ClearFormats
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]