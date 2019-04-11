---
title: ShowChartTipValues property (Excel Graph)
keywords: vbagr10.chm5207990
f1_keywords:
- vbagr10.chm5207990
ms.prod: excel
api_name:
- Excel.ShowChartTipValues
ms.assetid: aeff428a-01c2-51cc-2397-e178386e3e69
ms.date: 04/12/2019
localization_priority: Normal
---


# ShowChartTipValues property (Excel Graph)

**True** if charts show chart tip values. The default value is **True**. Read/write **Boolean**.


## Syntax

_expression_.**ShowChartTipValues**

_expression_ Required. An expression that returns one of the objects in the **Applies To** list.

## Example

This example turns off chart tip names and values.

```vb
With myChart.Application 
 .ShowChartTipNames = False 
 .ShowChartTipValues = False 
End With
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]