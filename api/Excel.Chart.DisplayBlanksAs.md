---
title: Chart.DisplayBlanksAs property (Excel)
keywords: vbaxl10.chm149101
f1_keywords:
- vbaxl10.chm149101
ms.prod: excel
api_name:
- Excel.Chart.DisplayBlanksAs
ms.assetid: b4e18939-6214-25e8-a0cd-c984b9f82346
ms.date: 04/16/2019
localization_priority: Normal
---


# Chart.DisplayBlanksAs property (Excel)

Returns or sets the way that blank cells are plotted on a chart. Can be one of the **[XlDisplayBlanksAs](Excel.XlDisplayBlanksAs.md)** constants. Read/write **Long**.


## Syntax

_expression_.**DisplayBlanksAs**

_expression_ A variable that represents a **[Chart](Excel.Chart(object).md)** object.


## Example

This example sets Microsoft Excel to not plot blank cells on Chart1.

```vb
Charts("Chart1").DisplayBlanksAs = xlNotPlotted
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]