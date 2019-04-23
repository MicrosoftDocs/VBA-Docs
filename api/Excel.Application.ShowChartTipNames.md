---
title: Application.ShowChartTipNames property (Excel)
keywords: vbaxl10.chm133208
f1_keywords:
- vbaxl10.chm133208
ms.prod: excel
api_name:
- Excel.Application.ShowChartTipNames
ms.assetid: 9f62fdc8-fcf0-eb4a-8ec4-d5d84cb96252
ms.date: 04/05/2019
localization_priority: Normal
---


# Application.ShowChartTipNames property (Excel)

**True** if charts show chart tip names. The default value is **True**. Read/write **Boolean**.


## Syntax

_expression_.**ShowChartTipNames**

_expression_ A variable that represents an **[Application](Excel.Application(object).md)** object.


## Example

This example turns off chart tip names and values.

```vb
With Application 
 .ShowChartTipNames = False 
 .ShowChartTipValues = False 
End With
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]