---
title: ShowChartTipNames property (Excel Graph)
keywords: vbagr10.chm5207986
f1_keywords:
- vbagr10.chm5207986
ms.prod: excel
api_name:
- Excel.ShowChartTipNames
ms.assetid: 0281bd54-2dbb-086f-23f7-ac507e19e519
ms.date: 04/12/2019
localization_priority: Normal
---


# ShowChartTipNames property (Excel Graph)

**True** if charts show chart tip names. The default value is **True**. Read/write **Boolean**.


## Syntax

_expression_.**ShowChartTipNames**

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